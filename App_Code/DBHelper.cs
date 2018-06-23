using System;
using System.Collections;  // List, ArrayList, HashTable
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;  // For Oracle and OleDbConnection
using System.Data.SqlClient;
using System.IO;
using System.Text;  // StringBuild, Encoding
using System.Web;  // HttpResponse
using System.Web.UI.WebControls; // DropDownList
using ClosedXML.Excel;
//using OfficeOpenXml;  // ExcelPackage. Need EPPlus.dll referenced


/// <summary>
/// This class makes it easier to run SQL. It can help produce an Excel "export" from any DataTable!
/// Example: var db = new DBHelper("sp_who2");
///          var csv = db.ExcelFormat(); // knows to use the DataTable from above and makes a string
///
/// You can get a single record like:
///    var db = new DBHelper("SELECT * FROM TABLE WHERE id=" + id);
///    var someField = db.Get("Name");
///
/// If you have update access to the table, you can update a table sooo easy...
///     var db = new DBHelper("SELECT * FROM table WHERE id = " + id);
///     db.Update("field", Field.Value);
///     db.Update("field2", Field2.Value);
///     db.Commit();
///
/// ADDing a record: var db = DBHelper.Insert("[TABLE NAME]");
///                  db.Set("field", Field.Value);
///                  db.Set("field2", Field2.Value);
///                  db.Commit();
///
/// It also has a .SQLtext property with the SQL including parameters which is nice for debugging
///     var param = new SqlParameter("@loginame", "active");
///     var db = new DBHelper("sp_who", param); // Yes you can also use sp_who @loginame='active'
///     string showSQL = db.SQLtext;
///
/// If you have multiple parameters you can use a collection and create it like:
///     var parameters = { new SqlParameter("@description", description), new SqlParameter("@data", data) };
///     var db = new DBHelper("SELECT * FROM Log WHERE description LIKE @description + '%' OR data LIKE @data + '%')", parameters);
/// </summary>
public class DBHelper
{
    /// <summary> Gets the DataTable inside  </summary>
    public DataTable DataTable = new DataTable();

    private DataSet ds = new DataSet();
    public SqlDataAdapter sda;
    public string SQLtext = "";
    public string TableName = "";
    public string IdentityColumn = "";
    public static int timeout = 10;
    public static string SecondLink = "";
    private SqlCommandBuilder scb;
    private static string dsn = null;
    public static bool AllowHTML = true;
    public DataRow CurrentRow;

    /// <summary>Constructor requires a DSN be defined in the app.config like:</summary>
    ///     <code><appSettings><add key="DSN" value="XXX"/></appSettings></code>
    public DBHelper()
    {
        SecondLink = "";
        // Rian: This gets called even for Static methods so DSN gets over-ridden
        
        if (dsn == null) dsn = HttpContext.Current.ApplicationInstance.Application["DSN"].ToString();
        //if (dsn == null) dsn = ConfigurationManager.AppSettings["DSN"]; // Reference System.Configuration
    }
    /// <summary> Create with a specific DataTable </summary>
    public DBHelper(DataTable myTable)
    {
        SecondLink = "";
        DataTable = myTable;
        //Rian: Now we use the one set Statically so it does not get over-written
        //dsn = ConfigurationManager.AppSettings["DSN"];
    }

    // Static Constructor for static methods to have DSN
    static DBHelper()
    {
        dsn = HttpContext.Current.ApplicationInstance.Application["DSN"].ToString();
        SecondLink = "";
        //if (String.IsNullOrEmpty(dsn)) throw new Exception("No DSN found in app.config");
    }

    /// <summary>
    /// Constructor that takes stored proc with no parameters
    /// </summary>
    /// <param name="sql">stored proc or statement</param>
    public DBHelper(string sql)
        : this() // Call default constructor (above with no paramaters)
    {
        connect(sql, dsn);
    }

    /// <summary>Runs the passed SQL against the connection styring passed</summary>
    /// <param name="sql"></param>
    /// <param name="dsn">Connection String</param>
    public DBHelper(string sql, string dsn)
        : this()
    {
        if (dsn == null || dsn == "") dsn = HttpContext.Current.ApplicationInstance.Application["DSN"].ToString();
        connect(sql, dsn);
    }

    /// <summary>Runs the passed SQL as parameterized query. Single SqlParameter is OK</summary>
    /// <param name="sql"></param>
    /// <param name="parameters">Array of SqlParameter objects</param>
    public DBHelper(string sql, params SqlParameter[] parameters)
        : this()
    {
        Connect(sql, dsn, parameters);
    }

    public DBHelper(string sql, string dsn, params SqlParameter[] parameters)
        : this()
    {
        Connect(sql, dsn, parameters);
    }

    /// <summary>Short-cut for the row count of the DataTable last run</summary>
    /// <returns>row count</returns>
    public int Length()  // We also have a .Count "property" this needs the ()
    {
        return DataTable.Rows.Count;
    }


    /// <summary>Pass 1 column name or CSV list. Short-cut for calling db.DataTable.Columns.Remove("columnName"); for each column to hide</summary>
    /// <param name="columnNames">Comma separated is OK</param>
    public void RemoveColumns(string columnNames)  // We also have a .Count "property" this needs the ()
    {
        this.DataTable.PrimaryKey = null; // In case they want to remove the first field
        foreach (var colName in columnNames.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries))
        {
            this.DataTable.Columns.Remove(colName);
        }
    }

    /// <summary>Converts a string OR object (that has a ToString) to a DateTime object and formats it like M/d/yyyy. </summary>
    ///	<param name="date">string or object that can be converted to a date</param>
    /// <returns> formatted string like 3/3/2015 </returns>
    public static string FormatDate(object date)
    {
        if (date == null || date.ToString() == "") return "";
        System.DateTime dt = System.DateTime.Parse(date.ToString());
        return dt.ToString("M/d/yyyy"); // 1 M = no leading zeros on month
    }

    public static string RelativeDate(object dateAsObject, bool utcTime=false)
    {
        if (dateAsObject.ToString() == "") return "";
        DateTime theDate = DateTime.Parse(dateAsObject.ToString());
        const int SecondsPerMinute = 60;
        const int SecondsPerHour   = 60 * SecondsPerMinute;
        const int SecondsPerDay    = 24 * SecondsPerHour;
        const int SecondsPerMonth  = 30 * SecondsPerDay;


        if (!utcTime) theDate = TimeZoneInfo.ConvertTimeToUtc(theDate, TimeZoneInfo.Local);
        var ts = new TimeSpan(DateTime.UtcNow.Ticks - theDate.Ticks);
        double delta = Math.Abs(ts.TotalSeconds);

        if (delta < 1 * SecondsPerMinute)    return ts.Seconds == 1 ? "one second ago" : ts.Seconds + " seconds ago";
        if (delta < 2 * SecondsPerMinute)    return "a minute ago";
        if (delta < 45 * SecondsPerMinute)   return ts.Minutes + " minutes ago";
        if (delta < 119 * SecondsPerMinute)  return "about an hour ago";
        if (delta < 24 * SecondsPerHour)     return ts.Hours + " hours ago";
        if (delta < 36 * SecondsPerHour)     return "yesterday";
        if (delta < 30 * SecondsPerDay)      return ts.Days + " days ago";
        if (delta < 12 * SecondsPerMonth)
        {
            int months = Convert.ToInt32(Math.Floor((double)ts.Days / 30));
            return months <= 1 ? "Over a month ago" : months + " months ago";
        }
        else
        {
            int years = Convert.ToInt32(Math.Floor((double)ts.Days / 365));
            return years <= 1 ? "Over a year ago" : "About " + years + " years ago";
        }
    }

    /// <summary>Converts a string or object with a date/time to HH:MM am/pm. </summary>
    /// <remarks> The example below shows you the native way to format it </remarks>
    ///	<param name="datetime">string or object that can be converted to a date</param>
    /// <returns> formatted time like 5:01 pm </returns>
    /// <example><code>
    /// System.DateTime dt = System.DateTime.Now;
    /// Response.Write(dt.ToString("h:mm tt"));
    /// </code></example>
    public static string FormatTime(object datetime)
    {
        if (datetime == null || datetime.ToString() == "") return "";
        System.DateTime dt = System.DateTime.Parse(datetime.ToString());
        return dt.ToString("h:mm tt");
    }

    /// <summary>Converts a string OR object (that has a ToString) to a DateTime object and formats it like M/d/yyyy. </summary>
    ///	<param name="date">string or object that can be converted to a date</param>
    /// <returns> formatted string like 3/3/2015 </returns>
    public static string FormatDateTime(object date)
    {
        if (date == null || date.ToString() == "") return "";
        System.DateTime dt = System.DateTime.Parse(date.ToString());
        return dt.ToString("M/d/yyyy h:mm tt"); // 1 M = no leading zeros on month
    }

    /// <summary> Writes to the Database in the LOG table </summary>
    public static void LogToDB(string description, string data, string hostname)
    {
        if (hostname == "") hostname = System.Net.Dns.GetHostName();
        SqlParameter[] parameters = {
					new SqlParameter("@description", description),
					new SqlParameter("@data", data),
					new SqlParameter("@computer", hostname) };
        new DBHelper("INSERT Into Log VALUES(@description, @data, @computer)", parameters);
    }

    /// <summary>Returns a DateTime column as a formatted date-only string</summary>
    /// <param name="field">Field name for column</param>
    /// <returns>Date formatted as M/d/yyyy</returns>
    public string AsDate(string field)
    {
        if (field == null || field == "") return "";
        string value = Get(field);
        if (value == "") return "";
        DateTime dt = DateTime.Parse(value);
        return dt.ToString("M/d/yyyy"); // 1 M = no leading zeros on month
    }
    /// <summary> This converts the value to h:mm am/pm formatted string </summary>
    public string AsTime(string field)
    {
        if (field == null || field == "") return "";
        string value = DataTable.Rows[0][field].ToString();
        if (value == "") return "";
        DateTime dt = DateTime.Parse(value);
        return dt.ToString("h:mm tt"); // 1 M = no leading zeros on month
    }

    /// <summary> Converts field value to a DateTime </summary>
    public DateTime GetDate(string field)
    {
        if (field == null || field == "") throw new Exception("No Field passed");
        string value = DataTable.Rows[0][field].ToString();
        if (value == "") throw new Exception("Blank Dates cannot be represented as a DateTime");
        return DateTime.Parse(value);
    }

    /// <summary> Converts field value to a bool </summary>
    public bool GetBool(string field)
    {
        string value = Get(field);
        return Convert.ToBoolean(value);
    }

    /// <summary>Adds a New Row to the .DataTable and returns a DBHelper instance</summary>
    /// <param name="tableName"></param>
    /// <param name="identityColumn"></param>
    /// <param name="database">Optional to change database away from PackageTechnology</param>
    /// <returns>DBHelper</returns>
    public static DBHelper Insert(string tableName, string identityColumn = "", string database = "PackageTechnology")
    {
        // This is the easiest way to get just the column/schema information

        DBHelper db = new DBHelper("select TOP 0 * from " + tableName, dsn.Replace("PackageTechnology", database)) { TableName = tableName };
        db.DataTable.TableName = tableName;
        //db.Columns[identityColumn].AutoIncrement = true;
        //db.Columns[identityColumn].Unique = true;
        if (identityColumn != "")
            db.IdentityColumn = identityColumn;
        else
            db.IdentityColumn = FindKeyField(tableName);

        db.AddRow();
        return db;
    }

    public static void Delete(string tableName, string deleteKey, string identityColumn = "")
    {
        if (identityColumn == "") identityColumn = FindKeyField(tableName);
        var db = new DBHelper("select TOP 1 * from " + tableName + " where " + identityColumn + "='" + deleteKey + "'") { IdentityColumn = identityColumn };
        if (db.DataTable.Rows.Count == 1)
        {
            db.DataTable.Rows[0].Delete();
            db.Commit();
        }
    }

    /// Note this is PRIVATE
    /// The constructors call this to not only make the connection but run the stored procedure with the parameters passed
    private void Connect(string sql, string dsn, params SqlParameter[] parameters)
    {
        SqlConnection connection = new SqlConnection(dsn);
        SqlCommand command = new SqlCommand(sql, connection);

        if (sql.StartsWith("sp")) command.CommandType = CommandType.StoredProcedure;
        command.CommandTimeout = timeout;
        if (parameters != null && parameters.Length > 0)
        {
            command.Parameters.AddRange(parameters);
        }

        BuildSQL(command, parameters); // Just to load .SQLText property to help user see/debug entire command
        sda = new SqlDataAdapter(command);
        sda.SelectCommand.CommandTimeout = timeout;
        sda.AcceptChangesDuringUpdate = true;
        sda.UpdateCommand = command;
        sda.FillLoadOption = LoadOption.PreserveChanges;
        try
        {
            sda.Fill(DataTable);
        }
        catch (Exception e)
        {
            throw new Exception(e.Message + " SQL:" + SQLtext);
        }
        if (DataTable.Rows.Count > 0)
            CurrentRow = DataTable.Rows[0];
    }

    // PRIVATE
    private void BuildSQL(SqlCommand command, SqlParameter[] parameters)
    {
        SQLtext = command.CommandText + " ";
        if (parameters != null && parameters.Length > 0)
        {
            foreach (SqlParameter p in parameters)
            {
                if (SQLtext.Contains(p.ParameterName))
                    SQLtext = SQLtext.Replace(p.ParameterName, "'" + p.Value.ToString() + "'");
                else
                    SQLtext += p.ParameterName + "='" + p.Value.ToString() + "', ";
            }
        }

        // Take off last comma
        if (SQLtext.Substring(SQLtext.Length - 2, 2) == ", ") SQLtext = SQLtext.Substring(0, SQLtext.Length - 2);
    }

    /// Note this is PRIVATE
    /// The constructors with a stored procedure but NOT parameters call this to connect to the database
    private void connect(string sql, string dsn)
    {
        SQLtext = sql;
        sda = new SqlDataAdapter(sql, dsn);
        sda.SelectCommand.CommandTimeout = timeout;
        sda.AcceptChangesDuringUpdate = true;
        var sqlUpper = sql.ToUpper();
        if (sqlUpper.StartsWith("SELECT") && !sqlUpper.Contains("JOIN"))
        {
            try
            {
                sda.FillSchema(DataTable, System.Data.SchemaType.Source); // This gets the column sizes (MaxLength) so the code can check at assignment for truncated strings
            }
            catch { }
        }

        sda.FillLoadOption = LoadOption.PreserveChanges;
        if (sql.StartsWith("sp") && !sql.Contains(" "))
        {
            sda.SelectCommand.CommandType = CommandType.StoredProcedure;
        }

        scb = new SqlCommandBuilder(sda);

        try // See: http://stackoverflow.com/questions/6696715/solving-a-timeout-error-for-sql-qurey
        {
            sda.Fill(DataTable);
        }
        catch // If a timeout or connection error then try at least 1 more time
        {
            try
            {
                sda.Fill(DataTable);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message + " SQL: " + sql + "<br>DSN: " + dsn);
            }
        }

        if (DataTable.Rows.Count > 0)
            CurrentRow = DataTable.Rows[0];
    }

    /// <summary>
    /// Runs stored procedures without parameters that return XML. This is notmally for when the XML is large and the SPROC uses FOR XML AUTO
    /// </summary>
    /// <param name="sproc">string</param>
    /// <returns>XML as a big string</returns>
    static public string ExecuteReturnXML(string sproc)
    {
        using (SqlConnection connection = new SqlConnection(dsn))
        using (SqlCommand command = new SqlCommand(sproc, connection))
        {
            command.CommandTimeout = timeout;
            connection.Open();

            System.Xml.XmlReader xmlr = command.ExecuteXmlReader();
            xmlr.Read();
            StringBuilder myString = new StringBuilder(4000);
            while (xmlr.ReadState != System.Xml.ReadState.EndOfFile)
            {
                myString.Append(xmlr.ReadOuterXml());
            }

            return myString.ToString();
        }
    }

    // Async call so it does not wait for results back from database. Usually for updates that take a while to make front-end faster
    // You cannot specify optional/default value for a parameter array so we must declare a separate overload method
    static public int ExecuteNonQuery(string sql)
    {
        return ExecuteNonQuery(sql, null); // Call over-load with 2 params
    }

    /// <summary>Takes SQL paramaterized fields and runs with .ExecuteNonQuery </summary>
    /// <param name="sql"></param>
    /// <param name="parameters"></param>
    static public int ExecuteNonQuery(string sql, params SqlParameter[] parameters)
    {
        using (SqlConnection connection = new SqlConnection(dsn))
        using (SqlCommand command = new SqlCommand(sql, connection))
        {
            if (sql.ToLower().StartsWith("exec")) command.CommandType = CommandType.StoredProcedure;
            if (sql.ToLower().StartsWith("sp")) command.CommandType = CommandType.StoredProcedure;
            command.CommandTimeout = timeout;
            if (parameters != null && parameters.Length > 0)
            {
                command.Parameters.AddRange(parameters);
            }

            connection.Open();
            return command.ExecuteNonQuery();
        }
    }

    /// <summary>Static version of the constructor. Timeout must be set through static timeout property</summary>
    /// <param name="sql"></param>
    /// <returns>DataTable</returns>
    static public DataTable ExecuteQuery(string sql)
    {
        SqlParameter[] empty = null;
        return ExecuteQuery(sql, empty); // Call over-load with 2 params
    }

    static public DataTable ExecuteQuery(string sql, string dsn) // Already Disconnected. No need for caller call .Close();
    {
        if (dsn.ToLower().Contains("driver"))
        {
            var oTable = new DataTable();
            OdbcConnection conn = new OdbcConnection(dsn);
            OdbcDataAdapter oDA = new OdbcDataAdapter(sql, conn);
            conn.Close();
            oDA.Fill(oTable);
            return oTable;
        }
        DataSet dataSet = new DataSet();
        OleDbConnection oConn = new OleDbConnection(dsn);
        OleDbDataAdapter adapter = new OleDbDataAdapter(sql, oConn);
        oConn.Close();
        adapter.Fill(dataSet, 0, 50000, "table");
        foreach (DataTable myTable in dataSet.Tables) // Loop just in-case first set is empty and the 2nd really has the results
        {
            if (myTable.Rows.Count > 0) return myTable;
        }
        return new DataTable(); // Return an empty Table if DataSet had none
    }

    /// <summary>Returns a DataTable given some SQL. If the SQL have @ parameters set the values in SqlParameter[]</summary>
    /// <param name="sql"></param>
    /// <param name="parameters"></param>
    /// <returns>DataTable</returns>
    static public DataTable ExecuteQuery(string sql, params SqlParameter[] parameters)
    {
        using (SqlConnection connection = new SqlConnection(dsn))
        using (SqlCommand command = new SqlCommand(sql, connection))
        {
            if (sql.ToLower().StartsWith("exec")) command.CommandType = CommandType.StoredProcedure;
            //if (sql.ToLower().StartsWith("sp")) command.CommandType = CommandType.StoredProcedure;
            command.CommandTimeout = timeout;
            if (parameters != null && parameters.Length > 0)
            {
                command.Parameters.AddRange(parameters);
            }

            connection.Open();
            SqlDataReader dr = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            return table;
        }
    }

    /// <summary>Returns HTML for 1 record suitable for a webpage output</summary>
    /// <returns>HTML Table with a class of stripe if further styling is required</returns>
    public string RecordToTable()
    {
        return RecordToTable("Database Record Details");
    }

    public string RecordToTable(string headerText)
    {
        // prints out the 1st record in a "field per row" table vs. a record per row table
        StringBuilder printHTML = new StringBuilder(1000);

        printHTML.Append("<table border='1' class='stripe' style='border: 1px solid black;'><thead>");
        printHTML.Append("<tr><th colspan='2'>" + headerText + "</th></tr></thead><tbody>");
        if (DataTable.Rows.Count > 0)
        {
            DataRow oRow = DataTable.Rows[0];
            foreach (DataColumn oCol in DataTable.Columns)
            {
                printHTML.Append("<tr><td class=prompt>" + oCol.ColumnName + "</td><td>");

                string type = oRow[oCol].GetType().ToString();
                if (type == "System.Decimal") // Align right and remove any .00
                    printHTML.Append(Convert.ToDouble(oRow[oCol].ToString()));
                else if (type == "System.Byte[]")
                    printHTML.Append(OIDToHex((byte[])oRow[oCol]));
                else if (type == "System.DateTime")
                    printHTML.Append(DBHelper.FormatDateTime(oRow[oCol].ToString()));
                else
                {
                    string val = oRow[oCol].ToString().Trim();
                    if (val == "?") val = "";
                    if (AllowHTML)
                        printHTML.Append(val);
                    else
                        printHTML.Append(val.Replace("<", "&lt;"));
                }

                printHTML.Append("</td><tr>");
            }
        }
        printHTML.Append("</tbody></table>");
        return printHTML.ToString();
    }

    /// <summary> Returns a HTML table with a class of stripe for styling.</summary>
    /// <returns>HTML Table as string</returns>
    public string MakeHTMLTable()
    {
        return MakeHTMLTable("");
    }

    /// <summary>Given a string of bytes (or field object from a DataRow) converts to hex representation starting with 0x</summary>
    /// <param name="OIDByteArray"></param>
    /// <returns></returns>
    public static string OIDToHex(object OIDByteArray)
    {
        if (OIDByteArray.GetType().ToString() != "System.Byte[]") return OIDByteArray.ToString().Trim();
        return OIDToHex((byte[])OIDByteArray);
    }

    // Over-load for above method
    public static string OIDToHex(byte[] OIDByteArray)
    {
        return "0x" + BitConverter.ToString(OIDByteArray).Replace("-", "");
    }

    /// <summary>Converts byte to hex (with leading 0 if needed). Called from OIDToHex</summary>
    /// <param name="data"></param>
    /// <returns>string</returns>
    public static string ToHex(int data)
    {
        return data.ToString("X2");
    }

    public static string ToHex(string data)
    {
        return ByteToHex(Encoding.ASCII.GetBytes(data));
    }

    public static string ByteToHex(byte[] ba)
    {
        string hex = BitConverter.ToString(ba);
        return hex.Replace("-", "");
    }

    public static string FromHex(string hexString)
    {
        int hexStringLength = hexString.Length;
        byte[] bytes = new byte[hexStringLength / 2];
        for (int i = 0; i < hexStringLength; i += 2)
        {
            int topChar = (hexString[i] > 0x40 ? hexString[i] - 0x37 : hexString[i] - 0x30) << 4;
            int bottomChar = hexString[i + 1] > 0x40 ? hexString[i + 1] - 0x37 : hexString[i + 1] - 0x30;
            bytes[i / 2] = (byte)(topChar + bottomChar);
        }
        return Encoding.ASCII.GetString(bytes);
    }

    /// <summary>Takes any object with a ToString and converts to a boolean</summary>
    /// <param name="o">Any string or object</param>
    /// <returns>bool</returns>
    public static bool ToBool(object o)
    {
        if (o == null) return false;
        if (o.ToString() == "") return false;
        if (o.ToString() == "1") return true; // brain dead overload that takes a string throws an error unless you pass "True"
        if (o.ToString() == "0") return false;// yes .Net says "0" is an error and not a boolean false
        return Convert.ToBoolean(o.ToString());
    }

    public static string ToDecimal(object str)
    {
        if (str.ToString() == "") return "";
        return Convert.ToDouble(str).ToString("G0");
    }

    public static double ToDouble(object str)
    {
        if (str.ToString() == "") return 0;
        return Convert.ToDouble(str);
    }


    /// <summary>
    /// Returns a HTML table with a class of stripe for styling.
    /// </summary>
    /// <param name="URL">optional. If 1st column should be clickable to this URL passing the value of the cell</param>
    /// <param name="openInNewTab">optional. set to true if link should open in a new tab/window</param>
    /// <returns>string</returns>
    public string MakeHTMLTable(string URL = "", bool hideFirstCol = false, int startrow = 1, bool openInNewTab = false, bool AllowHTML = true)
    {
        return DBHelper.MakeHTMLTable(this.DataTable, URL, hideFirstCol, startrow, openInNewTab);
    }

    /// <summary>Static version takes a DataTable</summary>
    /// <param name="db"></param>
    /// <param name="URL"></param>
    /// <param name="hideFirstCol"></param>
    /// <param name="startrow"></param>
    /// <param name="openInNewTab"></param>
    /// <returns>string</returns>
    public static string MakeHTMLTable(DataTable db, string URL = "", bool hideFirstCol = false, int startrow = 1, bool openInNewTab = false, bool AllowHTML = true)
    {
        StringBuilder printHTML = new StringBuilder(4000);
        printHTML.Append("<table class=\"stripe\"><thead><tr>");

        foreach (DataColumn oCol in db.Columns)
        {
            // Do we sometimes want to hide the 1st column? i.e. Make the next column in the result set show but use the ID in the link?
            if (hideFirstCol && oCol.Ordinal == 0) continue;
            printHTML.Append("<th>" + oCol.ColumnName + "</th>");
        }

        printHTML.Append("</tr></thead>\n<tbody>");
        int i = 0;
        foreach (DataRow oRow in db.Rows)
        {
            i++;
            if (i < startrow) continue; // Must have been on PREV page so skip

            printHTML.Append("<tr>");
            foreach (DataColumn oCol in db.Columns)
            {
                string url = URL;
                if (hideFirstCol && oCol.Ordinal == 0) continue; // Skip ID/OID column
                // If they passed us a URL then they want the 1st column clickable.
                if (URL != "" && (oCol.Ordinal == 0 || oCol.Ordinal == 1)) // They want this clickable
                {
                    int colToLink = oCol.Ordinal;
                    if (hideFirstCol) colToLink = 0;
                    if (!hideFirstCol && oCol.Ordinal == 1) // If NOT hidding 1st column then 2nd column should not be clickable unless baseURL is set
                    {
                    }
                    else
                    {
                        printHTML.Append("<td nowrap=1><a href=\"" + url + OIDToHex(oRow[colToLink]) + "\">" + OIDToHex(oRow[oCol.Ordinal]).Replace("<", "&lt;".Replace(">", "&gt;")) + "</a></td>");
                        continue;
                    }
                }

                if (SecondLink != "" && (oCol.Ordinal == 1 || oCol.Ordinal == 2)) // They want the 2nd column clickable clickable
                {
                    if (!hideFirstCol && oCol.Ordinal == 2) // If NOT hidding 1st column then 2nd column should not be clickable unless baseURL is set
                    {
                    }
                    else
                    {
                        printHTML.Append("<td nowrap=1><a href=\"" + SecondLink + OIDToHex(oRow[oCol.Ordinal]) + "\">" + OIDToHex(oRow[oCol.Ordinal]).Replace("<", "&lt;".Replace(">", "&gt;")) + "</a></td>");
                        continue;
                    }
                }

                string type = oRow[oCol].GetType().ToString();
                if (type == "System.Decimal") // Align right and remove any .00
                {
                    printHTML.Append("<td align=right>" + Convert.ToDouble(oRow[oCol].ToString()).ToString("G0") + "</td>");
                }
                else if (type == "System.Byte[]")
                {
                    printHTML.Append("<td>" + OIDToHex((byte[])oRow[oCol]) + "</td>");
                }
                else if (type == "System.DateTime")
                {
                    string date = DBHelper.FormatDate(oRow[oCol].ToString());
                    printHTML.Append("<td>" + date + "</td>");
                }
                else
                {
                    string val = oRow[oCol].ToString();
                    if (val == "?") val = ""; // Data Warehouses use a single ? like a NULL
                    if (AllowHTML)
                        printHTML.Append("<td>" + val + "</td>");
                    else
                        printHTML.Append("<td>" + RemoveTags(val) + "</td>");
                }
            }

            printHTML.Append("</tr>\n");
        }

        printHTML.Append("</tbody></table>");
        return printHTML.ToString();
    }

    /// <summary>Removes any HTML tags leaving just the contents or string portion.  Called from MakeHTMLTable </summary>
    /// <param name="s"></param>
    /// <returns>string</returns>
    public static string RemoveTags(string s)
    {
        if (s == "") return s;
        return System.Text.RegularExpressions.Regex.Replace(s, "\\<[^\\>]*\\>", " ").Trim();
    }

    /// <summary>Re-sorts the DataTable passed by the column passed and returns it</summary>
    /// <param name="oTable"></param>
    /// <param name="fieldname"></param>
    /// <returns>DataTable</returns>
    public static DataTable SortTable(DataTable oTable, string fieldname)
    {
        return FilterTable(oTable, null, fieldname);
    }

    /// <summary>Re-sort a DataTable in-memory</summary>
    /// <param name="fieldname"></param>
    public void SortTableInline(string fieldname)
    {
        FilterTableInline(null, fieldname);
    }

    /// <summary>Uses TABLE.Select() with an optional sort. Pass filter without a WHERE such as name='TEST'</summary>
    /// <param name="oTable"></param>
    /// <param name="filterExpression"></param>
    /// <param name="sortField"></param>
    /// <returns>New DataTable</returns>
    public static DataTable FilterTable(DataTable oTable, string filterExpression, string sortField = "")
    {
        System.Data.DataTable newTable = oTable.Clone(); // gets the meta data like field names
        foreach (DataRow oRow in oTable.Select(filterExpression, sortField)) // array of DataRows
            newTable.ImportRow(oRow); // add each new (filtered) row to new datatable
        return newTable;
    }

    /// <summary>Uses DataTable.Select() with an optional sort. Pass filter without a WHERE such as name='TEST'</summary>
    public DataTable FilterTable(string filterExpression, string sortField = "")
    {
        System.Data.DataTable newTable = DataTable.Clone(); // gets the meta data like field names
        foreach (DataRow oRow in DataTable.Select(filterExpression, sortField)) // array of DataRows
            newTable.ImportRow(oRow); // add each new (filtered) row to new datatable
        return newTable;
    }

    /// <summary>Returns part of a table based on a WHERE clause</summary>
    /// <param name="filterExpression"></param>
    /// <param name="sortField"></param>
    public void FilterTableInline(string filterExpression, string sortField = "")
    {
        System.Data.DataTable newTable = DataTable.Clone(); // gets the meta data like field names
        foreach (DataRow oRow in DataTable.Select(filterExpression, sortField)) // array of DataRows
            newTable.ImportRow(oRow); // add each new (filtered) row to new datatable
        DataTable = newTable;
    }

    public static DataTable Top(DataTable oTable, int count)
    {
        int i = 1;
        System.Data.DataTable newTable = oTable.Clone(); // gets the meta data like field names
        foreach (DataRow oRow in oTable.Rows) // array of DataRows
        {
            newTable.ImportRow(oRow); // add each new (filtered) row to new datatable
            if (i++ == count) break;
        }
        return newTable;
    }

    /// <summary>Returns a new table with only the top rows specified</summary>
    /// <param name="count">Integer</param>
    public void Top(int count)
    {
        int i = 1;
        DataTable newTable = DataTable.Clone(); // gets the meta data like field names
        foreach (DataRow oRow in DataTable.Rows) // array of DataRows
        {
            newTable.ImportRow(oRow); // add each new (filtered) row to new datatable
            if (i++ == count) break;
        }
        DataTable = newTable;
    }


    public static string DropDownOptions(DataTable oTable, string fieldName)
    {
        var db = new DBHelper();
        return db.DropDownOptions(oTable, fieldName, fieldName);
    }

    /// <summary>Returns a string of <option value="ZZZ">TEXT</option> tags. This overload takes a DataTable</summary>
    public string DropDownOptions(DataTable oTable, string fieldName, string valueName, string selectedValue = "")
    {
        this.DataTable = oTable;
        return DropDownOptions(fieldName, valueName, selectedValue);
    }

    /// <summary>Returns a string of <option value="ZZZ">TEXT</option> tags</summary>
    /// <param name="fieldName"></param>
    /// <param name="valueName"></param>
    /// <param name="selectedValue"></param>
    /// <returns>string</returns>
    public string DropDownOptions(string fieldName, string valueName, string selectedValue = "", string additionalOption = "")
    {
        StringBuilder outOptions = new StringBuilder(1000);
        if (fieldName == "" || fieldName == null) fieldName = "0"; //rs.Rows(0).name;
        if (IsNumeric(fieldName)) fieldName = DataTable.Columns[Convert.ToInt32(fieldName)].ColumnName;
        if (IsNumeric(valueName)) valueName = DataTable.Columns[Convert.ToInt32(valueName)].ColumnName;

        if (additionalOption != "") outOptions.Append("<option>" + additionalOption + "</option>");
        foreach (DataRow oRow in DataTable.Rows)
        {
            if (selectedValue == "")
                outOptions.Append("<option value=\"" + oRow[valueName] + "\">");
            else if (valueName != "")
                outOptions.Append("<option value=\"" + OIDToHex(oRow[valueName]).Replace("\"", "\"\"") + "\"" + IsSelected(selectedValue, OIDToHex(oRow[valueName])) + ">");
            else
                outOptions.Append("<option" + IsSelected(selectedValue, oRow[fieldName].ToString().Trim()) + ">");

            outOptions.Append(oRow[fieldName].ToString().Trim() + "</option>\n");
        }

        return outOptions.ToString();
    }

    /// <summary>All fields that match a parameter (a name= in the URL) will be set into the db Current Row BUT not committed</summary>
    public void FormToDB() // When a GET you can think of this as a split URL values and put them in matching fields
    {
        var db = this;
        foreach (DataColumn col in db.DataTable.Columns)
        {
            // Since Controls can't have spaces and reservced characters but fields can take these out!
            
            var controlName = col.ColumnName.Replace(" ", "").Replace("?", "").Replace("/", "").Replace("-", "").Replace("#", "").Replace("(", "").Replace(")", "").Replace("&", "");

            if (HttpContext.Current.Request.Form[controlName] == null && col.DataType.Name != "Boolean") continue; // This DB field was NOT passed as a form parameter (so do not blank it out)
            // When checkboxes are NOT checked they are NOT passed so we must assume this means 0
            var value = ""; // Init to blank for checkboxes that are not checked and thus do NOT exist (so they are null)
            if (HttpContext.Current.Request.Form[controlName] != null) value = HttpContext.Current.Request[controlName].ToString().Trim();
            if (col.DataType.Name == "Boolean") value = ToBool(value).ToString();
            db.Set(col.ColumnName, value);
        }
    }

    /// <summary>Builds a DropDownList for you from the DataTable. Requires Microsoft.Web</summary>
    /// <param name="controlID"></param>
    /// <param name="fieldName"></param>
    /// <param name="valueName"></param>
    /// <param name="selectedValue"></param>
    /// <returns>DropDownList</returns>
    public DropDownList CreateDropDownList(string controlID, string fieldName, string valueName = "", string selectedValue = "")
    {
        DropDownList ddl = new DropDownList();
        ddl.ID = controlID;

        foreach (DataRow oRow in DataTable.Rows)
        {
            if (valueName != "")
                ddl.Items.Add(new ListItem(oRow[fieldName].ToString(), oRow[valueName].ToString()));
            else
                ddl.Items.Add(new ListItem(oRow[fieldName].ToString()));
        }
        var foundItem = ddl.Items.FindByValue(selectedValue);
        if (foundItem != null) foundItem.Selected = true;
        return ddl;
    }

    public void LoadDropDown(string controlID, string fieldName, string selectedValue = "", string valueName = "")
    {
        var page = HttpContext.Current.Handler as System.Web.UI.Page;
        System.Web.UI.Control myControl = page.FindControl(controlID);
        var ddl = myControl as System.Web.UI.HtmlControls.HtmlSelect;

        foreach (DataRow oRow in DataTable.Rows)
        {
            if (valueName != "")
                ddl.Items.Add(new ListItem(oRow[fieldName].ToString(), oRow[valueName].ToString()));
            else
                ddl.Items.Add(new ListItem(oRow[fieldName].ToString()));
        }

        if (selectedValue != "")
        {
            var lookFor = ddl.Items.FindByValue(selectedValue);
            if (lookFor != null) lookFor.Selected = true;
        }
    }
 
    private string IsSelected(string str1, string str2)
    {
        if (str1.Trim().ToUpper() == str2.Trim().ToUpper())
            return " SELECTED";
        else
            return "";
    }

    /// <summary>
    /// Checks is the passed string is numeric.
    /// </summary>
    /// <param name="str"></param>
    /// <returns>boolean</returns>
    public static bool IsNumeric(object str)
    {
        decimal result;
        return Decimal.TryParse(str.ToString(), out result);
    }

    /// <summary>
    /// Short-cut for Regular Expression match. Returns true or fale
    /// </summary>
    /// <param name="str"></param>
    /// <param name="pattern"></param>
    /// <param name="IgnoreCase"></param>
    /// <returns>boolean</returns>
    public static bool IsMatchRegExp(string str, string pattern)
    {
        System.Text.RegularExpressions.Regex myRegx = new System.Text.RegularExpressions.Regex(pattern);
        return myRegx.IsMatch(str);
    }

    public static byte[] GetBytes(Stream theStream)
    {
        using (MemoryStream ms = new MemoryStream())
        {
            theStream.Seek(0, SeekOrigin.Begin);
            theStream.CopyTo(ms);
            return ms.ToArray();
        }
    }

    /// <summary>
    /// Builds a comma separated list and send it directly to the browser with the MIME type that tells the browser to launch Excel
    /// </summary>
    /// <param name="filename">optional. defaults to "Sheet1.csv"</param>
    public void StreamExcel(string filename = "Export.xlsx", bool wantHeadings = true, string tabName = "Sheet 1")
    {
        var stream = BuildExcel(wantHeadings, tabName);

        HttpContext.Current.Response.Clear();
        HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8;
        HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; // "application/vnd.ms-excel";
        HttpContext.Current.Response.AddHeader("Content-Disposition", "inline;filename=" + filename);
        HttpContext.Current.Response.BinaryWrite(GetBytes(stream));
        HttpContext.Current.Response.End();
    }

    public Stream BuildExcel(bool wantHeadings = true, string tabName = "Sheet 1")
    {
        var workbook = new XLWorkbook();
        workbook.Style.Font.FontName = "Calibri"; // Default Font for whole sheet
        workbook.Style.Font.FontSize = 11; // pixels in Excel or Point?

        var sheet = workbook.AddWorksheet(this.DataTable, tabName);
        //sheet.Row(1).CellsUsed().Style.Font.Bold = true;

        sheet.ColumnsUsed().AdjustToContents(); // Auto-Fit like double-clicking the right border (of the top A,B,C row)
        sheet.Row(1).Style.Fill.BackgroundColor = XLColor.CadetBlue;
        sheet.Row(1).Style.Font.FontSize = 10;

        IXLCells range = sheet.CellsUsed();
        var cellValue = "";
        foreach (IXLCell cellObj in range)
        {
            try
            {
                cellValue = cellObj.GetString();
            }
            catch //(Exception e)
            {
                cellValue = cellObj.ValueCached;
            }
            if (cellValue == null) break;

            if (cellValue.Contains("&#8805;")) cellObj.Value = cellValue.Replace("&#8805;", "≥");
            if (cellValue.Contains("&#8804;")) cellObj.Value = cellValue.Replace("&#8804;", "≤");
        }


        Stream theStream = new MemoryStream();
        workbook.SaveAs(theStream);
        theStream.Seek(0, SeekOrigin.Begin); // Set pointer to beginning of file/stream
        return theStream;
    }

    /*
    public ExcelPackage BuildExcel(bool wantHeadings = true, string tabName = "Sheet 1")
    {
        var p = new ExcelPackage();
        var ws = p.Workbook.Worksheets.Add(tabName);
        ws.Cells.Style.Font.Name = "Calibri"; // Default Font name for whole sheet
        ws.Cells.Style.Font.Size = 11;        // Default font size for whole sheet  
        int i = 1;
        if (wantHeadings)
        {
            ws.Cells[1, 1, 1, 100].Style.Font.Bold = true; // Whole 1st row is headings
            foreach (DataColumn theCol in DataTable.Columns)
            {
                ws.Cells[1, i++].Value = theCol.ColumnName;
            }
        }
        int row = 0;
        if (wantHeadings) row = 1;
        foreach (DataRow oRow in DataTable.Rows)
        {
            if (oRow.RowState == DataRowState.Deleted) continue;
            row++;
            int col = 1;
            foreach (DataColumn oCol in DataTable.Columns)
            {
                string type = oRow[oCol].GetType().ToString();
                if (type == "System.Decimal")
                    ws.Cells[row, col++].Value = Convert.ToDouble(oRow[oCol].ToString()); // Remove any .00
                else if (type == "System.DateTime")
                    ws.Cells[row, col++].Value = DBHelper.FormatDate(oRow[oCol].ToString());
                else // It is a string
                    ws.Cells[row, col++].Value = oRow[oCol].ToString().TrimEnd().Replace("&#8805;", "≥").Replace("&#8804;", "≤");
            }
        }
        ws.Cells[ws.Dimension.Address].AutoFitColumns(); // Autosize all columns
        return p;
    }
    */

    /// <summary>
    /// Builds a comma separated list and returns it as a string
    /// </summary>
    /// <param name="wantHeadings">bool</param>
    /// <param name="filename">optional. defaults to "Sheet1.csv"</param>
    public string ExcelFormat(bool wantHeadings = true, string filename = "Sheet1.csv")
    {
        string delimiter = ","; // TAB or "\t" does not work for CSV files
        StringBuilder returnString = new StringBuilder(4000);
        if (wantHeadings)
        {
            foreach (DataColumn oCol in DataTable.Columns)
            {
                string name = oCol.ColumnName;
                if (name == "ID") name = "id"; // See http://support.microsoft.com/kb/323626
                returnString.Append(name);
                if (oCol.Ordinal != DataTable.Columns.Count - 1) returnString.Append(delimiter); // If not last column add ,
            }
        }

        returnString.Append("\n");
        foreach (DataRow oRow in DataTable.Rows)
        {
            foreach (DataColumn oCol in DataTable.Columns)
            {
                string type = oRow[oCol].GetType().ToString();
                if (type == "System.Decimal")
                    returnString.Append(Convert.ToDouble(oRow[oCol].ToString())); // Remove any .00
                else if (type == "System.DateTime")
                    returnString.Append(DBHelper.FormatDate(oRow[oCol].ToString()));
                else // It is a string
                    returnString.Append(oRow[oCol].ToString().TrimEnd().Replace("\n", "").Replace("\r", "").Replace(delimiter, ";").Replace("\"", "'").Replace("&#8805;", "≥").Replace("&#8804;", "≤"));

                if (oCol.Ordinal != DataTable.Columns.Count - 1) returnString.Append(delimiter);
            }

            returnString.Append("\n");
        }
        return returnString.ToString().TrimEnd('\n');
    }


    /// <summary>Call when you want to Add a Row</summary>
    /// <returns>New DataRow</returns>
    public DataRow AddRow()
    {
        foreach (DataColumn c in DataTable.Columns)
        {
            c.AllowDBNull = true;
        }
        DataTable.Rows.Add(DataTable.NewRow());
        CurrentRow = DataTable.Rows[DataTable.Rows.Count - 1];
        sda.MissingSchemaAction = MissingSchemaAction.AddWithKey; // Must do this after the Add or it will already know which fields can't have NULL and immediately throw an exception
        return CurrentRow;
    }

    /// <summary>Gets the Count of records in the DataTable</summary>
    public int Count
    {
        get { return DataTable.Rows.Count; }
    }

    public string RowCount
    {
        get { return DataTable.Rows.Count.ToString(); }
    }

    /// <summary>Returns the .Rows of the DataTable</summary>
    public DataRowCollection Rows // Simulate the .Rows on a DataTable
    {
        get { return DataTable.Rows; }
    }

    /// <summary>>Returns the .Columns of the DataTable</summary>
    public DataColumnCollection Columns // Simulate the .Columns on a DataTable
    {
        get { return DataTable.Columns; }
    }

    /// <summary>Pass a fieldname and a row number to get the data</summary>
    /// <param name="key"></param>
    /// <param name="row"></param>
    /// <returns>string of the data</returns>
    public string Get(string key, int row)
    {
        if (DataTable.Rows.Count == 0) return "";
        key = key.Trim(); // Needed?
        if (key.Contains("OID"))
            return OIDToHex(DataTable.Rows[row][key]);
        else
        {
            var type = DataTable.Rows[row][key].GetType().ToString();
            if (type == "System.Decimal")
                return ToDecimal(DataTable.Rows[row][key]);
            else if (type == "System.DateTime")
                return FormatDate(DataTable.Rows[row][key]);
            else // Do we need System.Double???
                return DataTable.Rows[row][key].ToString().Trim();
        }
    }

    /// <summary>Get the passed field's data from the first row</summary>
    /// <param name="key"></param>
    /// <returns></returns>
    public string Get(string key)
    {
        // Rian: Changed from 0 to CurrentRow on 8/7/15 so we can use .Get("field") after setting CurrentRow
        if (CurrentRow == null)
            return Get(key, 0);
        else
        {
            int index = DataTable.Rows.IndexOf(CurrentRow);
            return Get(key, index);
        }
    }

    /// <summary>Get the passed field's data from the first row</summary>
    /// <param name="key"></param>
    /// <returns></returns>
    public string GetAsDate(string key)
    {
        if (this.Columns.Contains(key))
            if (this.Rows.Count == 0)
                return "";
            else
                return FormatDateTime(this.Rows[0][key].ToString());  // This calls above method which knows about CurrentRow cursor
        else
            return key + " not found";  // This calls above method which knows about CurrentRow cursor
    }

    public int GetAsInt(string key)
    {
        var str = Get(key);
        if (str == null || str == "") return 0;
        return Convert.ToInt32(str);
    }

    public string GetOID(string key)
    {
        return OIDToHex(CurrentRow[key]);
    }

    public static string Get(DataRow row, string key)
    {
        if (row == null) return "";
        if (key.Contains("OID"))
            return OIDToHex(row[key]);
        else
            return row[key].ToString().Trim();
    }

    public string Get(string key, DataRow row)
    {
        if (row == null) return "";
        if (key.Contains("OID"))
            return OIDToHex(row[key]);
        else
            return row[key].ToString().Trim();
    }

    /// <summary>Normally called internally to update a specific row and column</summary>
    public void Update(string key, object value, int row)
    {
        if (value.ToString() != "" && DataTable.Columns[key] != null)
        {
            if (DataTable.Columns[key].DataType.FullName == "System.Boolean") value = ToBool(value);
            if (DataTable.Columns[key].DataType.FullName.StartsWith("System.Int")) value = Convert.ToDecimal(value.ToString());
        }
        if (value.ToString() == "") value = DBNull.Value; // Many datatypes such as DateTime, int, float, etc. do NOT allow blanks
        DataTable.Rows[row][key] = value;
    }

    // This updated 1st row
    public void Update(string key, object value)
    {
        Update(key, value, 0);
    }

    /// <summary>Sets a field to the passed value for the CurrentRow. Often used after .AddRow() adds a new empty record </summary>
    /// <param name="key">field name</param>
    /// <param name="value">data to stick in the field</param>
    public void Set(string key, object value)
    {
        if (value == null || value.ToString() == "") value = DBNull.Value; // Many datatypes such as DateTime, int, float, etc. do NOT allow blanks
        CurrentRow[key] = value;
    }

    /// <summary>Call after Update to actually write the changes back to the database</summary>
    /// <returns>int with number of rows affected according to .Update() call</returns>
    public int Commit()
    {
        int rowsUpdated = 0;
        try
        {
            sda.AcceptChangesDuringUpdate = true;
            sda.AcceptChangesDuringFill = true;
            sda.MissingSchemaAction = MissingSchemaAction.AddWithKey;
            rowsUpdated = sda.Update(DataTable);
        }
        catch (SqlException e)
        {
            string err = "";
            foreach (DataRow row in DataTable.Rows)
            {
                if (row.RowState == DataRowState.Modified || row.RowState == DataRowState.Added)
                {
                    for (int i = 0; i < DataTable.Columns.Count; i++)
                    {
                        var t = row[i].GetType().ToString();
                        if (t == "System.String")
                        {
                            if (row[i].ToString().Length > DataTable.Columns[i].MaxLength)
                            {
                                err += "On row: " + i + " the value is too large for column: " + DataTable.Columns[i].ColumnName;
                                err += "\nThe data was: " + row[i].ToString();
                            }
                        }
                        if (row[i] == DBNull.Value)
                            err += "[" + DataTable.Columns[i].ColumnName + "]=null,\n";
                        else if (row[i].GetType().ToString() == "System.Boolean")
                            err += "[" + DataTable.Columns[i].ColumnName + "] = '" + Convert.ToBoolean(row[i]) + "',\n";
                        else if (row[i].GetType().ToString() == "System.Int32")
                            err += "[" + DataTable.Columns[i].ColumnName + "] = " + row[i] + ",\n";
                        else
                            err += "[" + DataTable.Columns[i].ColumnName + "] = '" + row[i] + "',\n";
                    }
                }
            }
            if (err.Length > 3) err = err.Substring(0, err.Length - 3);
            throw new Exception(e.Message + " " + err);
        }
        // If they are inserting 1 record re-read it for them
        if (TableName != "" && IdentityColumn != "")
        {
            DataTable = DBHelper.ExecuteQuery("SELECT * FROM " + TableName + " WHERE " + IdentityColumn + "=(SELECT MAX(" + IdentityColumn + ") FROM " + TableName + ")");
            CurrentRow = DataTable.Rows[0];

            if (IsNumeric(this.Get(IdentityColumn)))
                return Convert.ToInt32(this.Get(IdentityColumn));
            else
                return 0;
        }
        else
            return rowsUpdated;
    }

    /// <summary>This will delete the 2nd+ occurance of a row with a previously seen row for the column passed</summary>
    /// <param name="dTable">This is modified in-place because DataTable is a reference type</param>
    /// <param name="colName">Column name to look for dups in</param>
    public static void RemoveDuplicateRows(DataTable dTable, string colName)
    {
        Hashtable hTable = new Hashtable();
        ArrayList duplicateList = new ArrayList();

        foreach (DataRow drow in dTable.Rows)
        {
            if (hTable.Contains(drow[colName]))
                duplicateList.Add(drow);
            else
                hTable.Add(drow[colName], string.Empty);
        }

        foreach (DataRow dRow in duplicateList)
            dTable.Rows.Remove(dRow);
    }

    /// <summary>Util class now also contains this (since this is not really database specific)</summary>
    /// <param name="subject"></param>
    /// <param name="mailBody"></param>
    public static void EmailError(string subject, string mailBody)
    {
        string ErrorEmail = ConfigurationManager.AppSettings["errorEmail"];
        if (ErrorEmail == "" || ErrorEmail == null) ErrorEmail = "rianchipman@gmail.com"; // Fail-safe just in-case app.config is missing
        SendMail(ErrorEmail, "rianchipman@gmail.com", subject, mailBody, "");
    }

    public static void SendMail(string toList, string from, string subject, string msg, string cc = "")
    {
        toList = toList.Replace(";", ","); // Yes .net hates semicolons and throws an error
        System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(from, toList, subject, msg);
        message.IsBodyHtml = true;
        cc = cc.Replace(";", ",");
        foreach (string email in cc.Split(','))
        {
            if (email != "") message.CC.Add(new System.Net.Mail.MailAddress(email));
        }
        //message.Priority = System.Net.Mail.MailPriority.High;
        System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("mail.bigpond.com");
        try
        {
            client.Send(message);
        }
        catch
        {
            // Rian: Many errors say it could not connect but the email is sent fine. If 99% of the supposed
            // errors are not really errors that need to stop the whole program I say we just eat the error
            //throw new Exception("Error in Util.cs/SendMail(): " + e.Message + " " + e.StackTrace);
        }
    }

    /// <summary>Given a path to a .XLS or XLSX file returns a sheet as a DataTable</summary>
    /// <param name="excelFilePath">Full Path to the Excel file</param>
    /// <param name="sheetName">The name of the tab or sheet within the file</param>
    /// <param name="firstRowHeaders">Boolean. Pass false if the 1st row does not contain column headings</param>
    /// <returns>DataTable</returns>
    public static DataTable ExcelToDataTable(string excelFilePath, string sheetName = "", bool firstRowHeaders = true)
    {
        string headers = "NO";
        if (firstRowHeaders) headers = "YES";

        //string dsn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 8.0;HDR=" + headers + "\"";
        string dsn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 8.0;HDR=" + headers + "\"";

        if (excelFilePath.ToLower().Contains(".xlsx") || excelFilePath.ToLower().Contains(".xlsm"))
        {
            dsn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=" + headers + "\"";
        }
        OleDbConnection dbConnection = new OleDbConnection(dsn);
        dbConnection.Open();

        if (sheetName == "") sheetName = GetFirstSheetName(dbConnection);
        if (!sheetName.EndsWith("$")) sheetName += "$";

        string sql = "SELECT * FROM [" + sheetName + "]"; // Name of "tab" or sheet. Default is Sheet1$

        DataTable dt = new DataTable(sheetName); // Name the table the same as the sheet
        OleDbDataAdapter dbAdapter = new OleDbDataAdapter(sql, dbConnection);
        dbAdapter.Fill(dt);
        dbConnection.Close();
        return dt;
    }

    /// <summary>Returns the name of the first sheet/tab regardless if it is hidden of not </summary>
    /// <param name="dbConnection"></param>
    /// <returns>string</returns>
    public static string GetFirstSheetName(OleDbConnection dbConnection)
    {
        DataTable dbSchema = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        if (dbSchema == null || dbSchema.Rows.Count < 1) throw new Exception("No Sheet found in .XLS file.");
        string sheetName = dbSchema.Rows[0]["TABLE_NAME"].ToString();
        //dbConnection.Close(); // They probably do not want us closing this connection because with OLEDB if they have it open we close it
        return sheetName.Replace("'", "");
    }

    /// <summary>
    /// Adds a filename to a database field (normally a varbinary(MAX)) as a a stream of bytes
    /// This is raw but in SQL Server Management Studio it will appear as a Hex string but this is just their representation
    /// </summary>
    /// <param name="filename">@"C:\TEMP\file.jpg"</param>
    /// <param name="sqlWithAtFile">"UPDATE DesignRule SET Image = @File WHERE RuleID=" + ruleID;</param>
    public static void FileToColumn(string filename, string sqlWithAtFile)
    {
        var stream = new FileStream(filename, FileMode.Open, FileAccess.Read);
        var reader = new BinaryReader(stream);
        byte[] file = reader.ReadBytes((int)stream.Length);
        SqlConnection connection = new SqlConnection(dsn);
        connection.Open(); // Why use using () when it's going out of scope in 2 more lines anyway
        var sqlWrite = new SqlCommand(sqlWithAtFile, connection);
        sqlWrite.Parameters.Add("@File", SqlDbType.VarBinary, file.Length).Value = file;
        sqlWrite.ExecuteNonQuery();
        reader.Close();
        stream.Close();
        reader.Dispose();
        stream.Dispose();
    }

    /// <summary>
    /// Reads a binary stream from a column (normally a varbinary(MAX)) and writes it to a filename (and path)
    /// </summary>
    /// <param name="filename">@"C:\TEMP\file.jpg"</param>
    /// <param name="sqlToReturnImage">SELECT Image FROM DesignRule WHERE RuleID=" + ruleID</param>
    public static void FileFromColumn(string filename, string sqlToReturnImage)
    {
        MemoryStream stream = FileRead(sqlToReturnImage);
        SaveImage(stream, filename);
    }

    /// <summary>
    /// Given SQL such as SELECT Image FROM DesignRule WHERE RuleID=1 returns a MemoryStream object
    /// </summary>
    /// <param name="sqlToReturnImage">SELECT Image FROM DesignRule WHERE RuleID=" + ruleID</param>
    /// <returns>MemoryStream</returns>
    public static MemoryStream FileRead(string sqlToReturnImage)
    {
        MemoryStream memoryStream = new MemoryStream();
        SqlConnection connection = new SqlConnection(dsn);
        connection.Open();
        var sqlQuery = new SqlCommand(sqlToReturnImage, connection);
        var sqlQueryResult = sqlQuery.ExecuteReader();
        if (sqlQueryResult == null) return null;
        sqlQueryResult.Read();
        var blob = new Byte[(sqlQueryResult.GetBytes(0, 0, null, 0, int.MaxValue))];
        sqlQueryResult.GetBytes(0, 0, blob, 0, blob.Length);
        memoryStream.Write(blob, 0, blob.Length);

        return memoryStream;
    }

    //public static System.Drawing.Image StreamToImage(MemoryStream stream)
    //{
    //    return System.Drawing.Image.FromStream(stream);
    //}

    /// <summary>
    /// Saves a MemoryStream to a file. This can be thought of serializing to a JPG
    /// </summary>
    /// <param name="stream">Any MemoryStream object</param>
    /// <param name="filename">Filename including path</param>
    public static void SaveImage(MemoryStream stream, string filename)
    {
        using (FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write))
        {
            stream.WriteTo(file);
            file.Close();
        }
        stream.Close();
    }

    /// <summary>Finds the Primary key by running sp_pkeys on the table passed or assuming the first column if no key is officially set </summary>
    /// <param name="tableName"></param>
    /// <returns>Name of column</returns>
    public static string FindKeyField(string tableName, string dsn = "")
    {
        string sql = "sp_pkeys '" + tableName.Replace("dbo.", "") + "'"; // Just the table portion of dbo.table
        if (dsn == "") dsn = ConfigurationManager.AppSettings["DSN"];
        DBHelper db = new DBHelper(sql, dsn);
        if (db.DataTable.Rows.Count > 0)
            return db.Get("COLUMN_NAME");

        sql = "SELECT TOP 0 * FROM " + tableName; // Just to get column names
        db = new DBHelper(sql, dsn);
        return db.DataTable.Columns[0].ColumnName;
    }

    public static string UrlEncode(object str)
    {
        return HttpUtility.UrlEncode(str.ToString().Trim());
    }

    public static string UrlDecode(object str)
    {
        return HttpUtility.UrlDecode(str.ToString().Trim());
    }

    public static string HtmlEncode(object str)
    {
        return HttpUtility.HtmlEncode(str.ToString().Trim());
    }

    public static string HtmlDecode(object str)
    {
        return HttpUtility.HtmlDecode(str.ToString().Trim());
    }
} // end Class

/// <summary>Creates a .SetColumnsOrder method on DataTable objects
/// USAGE: table.SetColumnsOrder("Qty", "Unit", "Id");
/// </summary>
public static class DataTableExtensions
{
    // The this DataTable is NOT a real passed parameter. As an extension this is auto-populated
    public static void SetColumnsOrder(this DataTable table, string columnNames)
    {
        var columnIndex = 0;
        foreach (var columnName in columnNames.Split(','))
            table.Columns[columnName].SetOrdinal(columnIndex++);
    }
}
