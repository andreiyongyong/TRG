using System;					// for Convert, StringComparision, etc.
using System.Collections.Generic;// for <List>
using System.Configuration;     // For reading app.config
using System.Diagnostics;		// for EventLog and EventLogEntryType
using System.IO;				// for File and Path classes
using System.Linq;              // for .Where support
using System.Net;				// for WebClient and CredentialClass
using System.Security.Principal;// For WindowsIdentity
using System.Text;				// for StringBuilder and UTF8Encoding
using System.Web;				// for HttpContext
using System.Xml;				// for XmlDocument, XmlTextWritter, XmlNodeNode
//using System.DirectoryServices.AccountManagement; // Need System.DirectoryServices.AccountManagement.dll referenced
using System.Data;
using System.Web.UI.HtmlControls; // Need System.Web.dll referenced
using System.Web.UI.WebControls;  // For all HtmlInput type classes
using System.Drawing.Imaging; // EncodeInfo
using System.Drawing;

/*****************************************************************
SAMPLE CALLS:

    var contents = Util.ReadFile(filename);
    Util.AppendFile(filename, DateTime.Now + " Hello World");
or
    bool isInt = Util.IsInteger("no");
*****************************************************************/

namespace System
{
    /// <summary> Common Web-Related functions  </summary>
    public class Util
    {
        public static int maxRetries = 3; // Used for File IO retries (if log file is locked, etc.)
        public static int retryMilliSeconds = 1000; // Used for sleep time if a file is locked
        public static string ErrorEmail = "";
        public static string SmtpServer = "mail.bigpond.com"; // Default to SMTP alias (load-balanced)
        public static string ProxyServer = "proxy"; // Default to same one IE uses
        public static bool CaseSensitive = false;
        private static StringComparison ignoreCase = StringComparison.OrdinalIgnoreCase;

        /// <summary> Default constructor not needed </summary>
        public Util() // Default constructor
        {
            string ErrorEmail = ConfigurationManager.AppSettings["errorEmail"];
            if (ErrorEmail == "") ErrorEmail = "rianchipman@gmail.com"; // Failsafe just in-case!

            SmtpServer = ConfigurationManager.AppSettings["mailServer"];
            if (SmtpServer == null) SmtpServer = "mail.bigpond.com"; // Failsafe just in-case!
        }

        /// <summary> Capitalizes the 1st character in each word. Acronymn exceptions can be added inside</summary>
        public static string TitleCase(string str)
        {
            if (str == null) str = "";
            str = str.ToLower();
            str = str.Replace("dram", "DRAM").Replace("ram", "RAM").Replace("nand", "NAND").Replace("ddr", "DDR").Replace("cmos", "CMOS").Replace("pt ", "PT ").Replace("kdg ", "KDG ").Replace("imft ", "IMFT ");

            System.Globalization.TextInfo myTI = new System.Globalization.CultureInfo("en-US", false).TextInfo;
            return myTI.ToTitleCase(str);
        }

        public static void BuildDropDown(string selectName, DataTable oTable, string fieldName, string selectedValue, string valueName)
        {
            BuildDropDown(selectName, oTable, fieldName, selectedValue, valueName, "", "");
        }

        public static void BuildDropDown(string selectName, DataTable oTable, string fieldName, string selectedValue, string valueName, string firstOptionDisplay, string selectTagAttr)
        {
            if (fieldName == "" || fieldName == null) fieldName = "0"; //rs.Rows(0).name;
            HttpContext.Current.Response.Write("<select name='" + selectName + "'");
            if (selectTagAttr != "") HttpContext.Current.Response.Write(" " + selectTagAttr);
            HttpContext.Current.Response.Write(">");
            if (firstOptionDisplay != "") Print("<option value=''>" + firstOptionDisplay + "</option>");
            foreach (DataRow oRow in oTable.Rows)
            {
                if (valueName != "")
                    Print("<option value=\"" + oRow[valueName].ToString().Trim() + "\"" + IsSelected(selectedValue, oRow[valueName].ToString().Trim()) + ">");
                else
                {
                    HttpContext.Current.Response.Write("<option");
                    if (selectedValue != "") HttpContext.Current.Response.Write(IsSelected(selectedValue, oRow[valueName].ToString().Trim()));
                    HttpContext.Current.Response.Write(">");
                }
                Print(oRow[fieldName].ToString().Trim() + "</option>");
            }
            Print("</select>");
        }

        /// <summary>Returns DataTable from CSV file. 1st line must have headings</summary>
        /// <param name="fileWithPath">string</param>
        /// <returns>DataTable</returns>
        public static DataTable CSVFileToDataTable(string fileWithPath, bool firstLineContainsHeadings = true)
        {
            var fileOnly = Path.GetFileName(fileWithPath);
            var dirOnly = Path.GetDirectoryName(fileWithPath);
            var dsn = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" + dirOnly + ";HDR=YES;Extensions=asc,csv,tab,txt;Persist Security Info=False";

            var sql = "select * from [" + fileOnly + "]";
            var theTable = new DataTable();
            var conn = new System.Data.Odbc.OdbcConnection(dsn);
            var oDA = new System.Data.Odbc.OdbcDataAdapter(sql, conn);
            conn.Close();
            oDA.Fill(theTable);
            if (firstLineContainsHeadings == false)
            {
                var theRow = theTable.NewRow();
                for (int i = 0; i < theTable.Columns.Count; i++)
                {
                    theRow[i] = theTable.Columns[i].ColumnName;
                    theTable.Columns[i].ColumnName = "Column" + (i+1);
                }

                theTable.Rows.Add(theRow);
            }
            foreach (DataColumn col in theTable.Columns) col.ColumnName = col.ColumnName.Trim(); // Just in-case the have leading or trailing spaces
            return theTable;
        }

        /// <summary> Reads a comma seperated string and returns a table using Microsoft ODBC Text Driver</summary>
        public static DataTable CSVtoDataTable(string csv, bool firstLineContainsHeadings = true)
        {
            if (!firstLineContainsHeadings)
            {
                var topLine = csv.Split(new string[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries)[0]; // Only get top line!
                var columns = topLine.Split(',');
                var headerLine = "";
                for (int j = 1; j <= columns.Length; j++) headerLine += "Column" + j.ToString() + ",";
                headerLine = headerLine.TrimEnd(',');
                csv = headerLine + "\n" + csv;
            }
            var file = Util.GetEnv("TEMP") + @"\temp.csv";
            Util.WriteFile(file, csv); // The "trick" to using Microsoft's code/driver to parse CSV chunks is to write it to disk so we use wherever their %TEMP% points to
            return CSVFileToDataTable(file);
        }

        /// <summary> If any controls match the name of a field (excluding spaces) then load the value. Similar to DefaultInputs</summary>
        /// <param name="db">DBHelper</param>
        // Loop over all columns in a record & Look for a control named the exact same thing as the field
        // (replacing spaces with nothing) and If found set the .Value or .Text property to the DB value
        public static void SetControlsFromDB(DBHelper db, bool splitLinkFields = false)
        {
            if (db.Count == 0) return;
            DataRow row = db.Rows[0]; // 1st (and only) row
            var page = HttpContext.Current.Handler as System.Web.UI.Page;
            foreach (DataColumn col in db.DataTable.Columns)
            {
                var controlName = col.ColumnName.Replace(" ", "").Replace("?", "").Replace("/", "").Replace("-", "").Replace("#", "").Replace("(", "").Replace(")", "").Replace("&", "");
                System.Web.UI.Control myControl = page.FindControl(controlName);
                if (myControl == null) continue; // SKIP
                var value = db.Get(col.ColumnName);
                var type = myControl.GetType().ToString();
                if (myControl is HtmlInputText)
                {
                    if (col.ColumnName.Contains("Link") || col.ColumnName == "Preliminary Design Internal Review Feedback" || col.ColumnName == "Preliminary Design External Review Feedback" || col.ColumnName == "Stack Analysis Feedback (HVM site)" || col.ColumnName == "Wire Fanout Feedback (HVM Site)")
                    {
                        (myControl as HtmlInputText).Value = Util.LeftOf(value, ",", true);
                        myControl = page.FindControl(controlName + "Desc");
                        if (myControl != null) (myControl as HtmlInputText).Value = Util.RightOf(value, ",", true);
                    }
                    else
                        (myControl as HtmlInputText).Value = value;
                }
                else if (myControl is HtmlInputHidden)
                    (myControl as HtmlInputHidden).Value = value;
                else if (myControl is HtmlSelect)
                    (myControl as HtmlSelect).Value = value;
                else if (myControl is HtmlTextArea)
                    (myControl as HtmlTextArea).Value = value;
                else if (myControl is HtmlInputCheckBox && value != "" && value != "False")
                    (myControl as HtmlInputCheckBox).Checked = true;
                else if (myControl is CheckBoxList && value != "")
                {
                    foreach (var item in value.Split(','))
                    {
                        var checkbox = (myControl as CheckBoxList).Items.FindByText(item);
                        if (checkbox != null) checkbox.Selected = true;
                    }
                }
                else if (myControl is RadioButtonList && value != "")
                {
                    var radioButton = (myControl as RadioButtonList).Items.FindByText(value);
                    if (radioButton != null) radioButton.Selected = true;
                }
                else if (myControl is Label || myControl is Literal)
                {
                    if (col.ColumnName == "Package Designer" || col.ColumnName == "Assembly Central Engineer" || col.ColumnName == "Product Engineer (NVE/DEG)" || col.ColumnName == "Simulation Engineer")
                    {
                        var newValue = "";
                        foreach (string user in value.Split(','))
                            newValue += "<a href='EmpDetails.aspx?username=" + Between(user, "(", ")") + "'>" + user + "</a>, ";
                        if (value.Trim() != "") value = "<div class=rowline><b>" + col.ColumnName + "</b> " + newValue.TrimEnd(' ').TrimEnd(',') + "</div>";
                    }
                    else if (col.ColumnName.Contains("Link"))
                    {
                        var link = value;
                        value = Util.RightOf(link, ",", true);
                        link = Util.LeftOf(link, ",", true);
                        if (link != "") value = "<div class=rowline><b>" + col.ColumnName + "</b> <a href='" + link + "'>" + value.Replace("%20", " ") + "</a></div>";
                    }
                    else if (col.ColumnName == "DesignerName" || col.ColumnName == "DrawnName") value = "<a href='Emp.aspx?lastname=" + value + "'>" + value + "</a>";
                    else if (col.ColumnName.ToLower().Contains("username")) value = "<a href='EmpDetails.aspx?username=" + value + "'>" + value + "</a>";
                    else if (col.ColumnName.Contains("DRN") && value.Length == 12) value = "DRN# <a href='ADRDisplay.aspx?ADR=ADR" + value + "'>" + value + "</a>";
                    else if (col.ColumnName.Contains("DCN") && value.Length == 12) value = "DCN# <a href='ADRDisplay.aspx?ADR=ADR" + value + "'>" + value + "</a>";
                    // Now that any special values have been set it into the Control's Text property
                    if (myControl is Literal)
                        (myControl as Literal).Text = value;
                    else
                        (myControl as Label).Text = value;
                }
            }
        }

        /// <summary> runat=server inputs and drop-downs will be set to matching URL Request named values </summary>
        public static void DefaultInputs()
        { // Loop over each parameter value passed and set the value back to what they entered
            var page = HttpContext.Current.Handler as System.Web.UI.Page;

            foreach (string theName in HttpContext.Current.Request.QueryString)
            {
                var value = HttpContext.Current.Request.QueryString[theName];
                if (value == null || value == "") continue;
                try
                {
                    System.Web.UI.Control myControl = page.FindControl(theName);
                    if (myControl == null) continue;
                    if (myControl is System.Web.UI.HtmlControls.HtmlInputText)
                        (myControl as System.Web.UI.HtmlControls.HtmlInputText).Value = value;
                    else if (myControl is System.Web.UI.HtmlControls.HtmlSelect)
                        (myControl as System.Web.UI.HtmlControls.HtmlSelect).Value = value;
                }
                catch { }
            }

            foreach (string theName in HttpContext.Current.Request.Form)
            {
                var currentControl = new System.Web.UI.Control();
                var value = HttpContext.Current.Request.Form[theName];
                if (value == null || value == "") continue;
                value = value.Trim();
                System.Web.UI.Control myControl = page.FindControl(theName);
                if (myControl == null) continue;
                if (myControl is System.Web.UI.HtmlControls.HtmlInputText)
                    (myControl as System.Web.UI.HtmlControls.HtmlInputText).Value = value;
                else if (myControl is System.Web.UI.HtmlControls.HtmlSelect)
                    (myControl as System.Web.UI.HtmlControls.HtmlSelect).Value = value;
            }
        }


        /// <summary> Sends an email to the errorEmail in the web.config or set with the .ErrorEmail property. </summary>
        public static void EmailError(string subject, string mailBody)
        {
            string ErrorEmail = ConfigurationManager.AppSettings["errorEmail"];
            if (ErrorEmail == "" || ErrorEmail == null) ErrorEmail = "rchipman@micron.com"; // Fail-safe just in-case app.config is missing

            SendMail(ErrorEmail, "asmpd@micron.com", subject, mailBody);
        }

        /// <summary> Returns true if string starts with a number </summary>
        public static bool StartsNumber(string str)
        {
            if (String.IsNullOrEmpty(str)) return false;
            char c = str[0];
            return c >= 48 && c <= 57; // digits are in the range [48,57] in ASCII code
        }

        /// <summary> Returns false if string starts with a number </summary>
        public static bool StartsLetter(string str)
        {
            return !StartsNumber(str);
        }

        /// <summary> Tries to return a DateTime onject given a string in various formats </summary>
        public static DateTime CleanDate(object badDate)
        {
            if (badDate == null) return new DateTime();
            var fixDate = badDate.ToString();
            if (fixDate == "") return new DateTime();
            fixDate = fixDate.Replace("st", "").Replace("nd", "").Replace("rd", "").Replace("th", ""); // Spaces are OK. .NET's .parse must trim them out
            fixDate = fixDate.Replace("after ", "").Replace("After ", "").Replace("AFTER ", "").Replace("before ", "").Replace("(ES)", "").Replace("(EST)", "").Replace("(Est)", "").Replace("(sim)", "");
            if (!Util.IsDate(fixDate))
            {
                if (StartsLetter(fixDate)) // Jul-31
                {
                    // 1st we split up the parts with either a dash or a slash and trim out any blanks
                    var dateParts = fixDate.Split(new string[] { "-", "/" }, StringSplitOptions.RemoveEmptyEntries);
                    if (dateParts.Length == 3)
                        fixDate = dateParts[1] + "-" + Util.Left(dateParts[0], 3) + "-" + dateParts[2];
                    else if (dateParts.Length == 2) // July-31
                        fixDate = dateParts[1] + "-" + Util.Left(dateParts[0], 3) + "-" + DateTime.Now.Year; // Add 2015
                }
                else //
                {
                    var dateParts = fixDate.Split(new string[] { "-", "/" }, StringSplitOptions.RemoveEmptyEntries);
                    if (dateParts.Length == 3)
                        fixDate = dateParts[0] + "-" + Util.Left(dateParts[1], 3) + "-" + dateParts[2];
                    else if (dateParts.Length == 2) // 31-Jul
                        fixDate = dateParts[0] + "-" + Util.Left(dateParts[1], 3) + "-" + DateTime.Now.Year; // Add 2015

                    if (!Util.IsDate(fixDate) && dateParts.Length == 3)
                    {
                        fixDate = Util.Left(dateParts[1], 3) + "-" + dateParts[0] + "-" + dateParts[2]; // Add 2015
                    }
                }
                if (!Util.IsDate(fixDate)) return new DateTime(); // Still bad=create base date (non-nullable datatype so we can;t just pass back nothing)
            }
            DateTime retDate = DateTime.Parse(fixDate);

            if (Util.Contains(badDate.ToString(), "after")) retDate = retDate.AddMonths(1).AddDays(-1);
            return retDate;
        }

        /// <summary>Returns the username from the SiteMinder cookie token</summary>
        /// <returns>username</returns>
        //public static string SMGetUsername(string smtoken)
        //{
        //   if (smtoken == "") return "";
        //    StringBuilder username = new StringBuilder(13);
        //    c_SMGetUsername(smtoken, username, 13);
        //    return username.ToString();
        //}


        /// <summary> Replace all tags with nothing </summary>
        public static string StripHtml(object htmlObj)
        {
            var html = htmlObj.ToString();
            return System.Text.RegularExpressions.Regex.Replace(html, "<.*?>", "").Trim();
        }

        /// <summary> Returns full name given a username </summary>
        public static string GetNameFromUsername(string username)
        {
            var sql = "SELECT first_name + ' ' + last_name as name FROM SAP_worker WHERE micron_username=@username";
            var parameter = new System.Data.SqlClient.SqlParameter("@username", username);
            var db = new DBHelper(sql, ConfigurationManager.AppSettings["PT"], parameter);
            if (db.Count == 0) return null;
            return db.Get("name");
        }

        /// <summary> Returns worker # given a username </summary>
        public static int GetEmpNoFromUsername(string username) // INTs go up to 2,147,483,647
        {
            var sql = "SELECT worker_no FROM SAP_worker emp WHERE micron_username=@username";
            var parameter = new System.Data.SqlClient.SqlParameter("@username", username);
            var db = new DBHelper(sql, ConfigurationManager.AppSettings["PT"], parameter);
            if (db.Count == 0) return -1;
            return Util.ToInt(db.Get("worker_no"));
        }

        /// <summary> Reurns the current logged in username </summary>
        public static string Username // Property accessor for .GetUsername() below
        {
            get
            {
                return GetUsername();
            }
        }

        /// <summary>Returns the username from either the authenticated user or SiteMinder or the remote host</summary>
        /// <returns>username</returns>
        public static string GetUsername()
        {
            if (HttpContext.Current != null) // Running via a web-server context and not a windows application
            {
                string user = HttpContext.Current.Request.ServerVariables["AUTH_USER"]; // WINNTDOM\rchipman
                user = RightOf(user, @"\");
                if (user != "") return user;
            }

            string username = GetCookie("DCTMLOGIN");
            if (username != "") return username;

            //string token = GetSMToken();
            //if (token != "") username = SMGetUsername(token);
            //if (username != "") return username;

            try
            {
                if (HttpContext.Current == null) // Running via a web-server context and not a windows application
                {
                    if (WindowsIdentity.GetCurrent() != null)
                    {
                        string domainAndUser = WindowsIdentity.GetCurrent().Name;
                        return RightOf(domainAndUser, @"\");
                    }
                    return Environment.GetEnvironmentVariable("USERNAME");
                }
                else
                {
                    string host = Dns.GetHostName();
                    if (HttpContext.Current != null) // Running via a web-server context and not a windows application
                    {
                        host = HttpContext.Current.Request.ServerVariables["REMOTE_HOST"]; // Maybe IP address
                    }
                    host = Dns.GetHostEntry(host).HostName; // Make sure we have the hostname now
                    host = host.Substring(0, host.IndexOf("."));
                    if (IsNumeric(host.Substring(host.Length - 1, 1))) host = host.Substring(0, host.Length - 1);
                    return host;
                }
            }
            catch
            {
                return Environment.GetEnvironmentVariable("USERNAME");
            }
        }
        /*
        public static string GetSMSession(string username, string password)
        {
            StringBuilder token = new StringBuilder(2048);
            int len = 2048;
            int stat = c_SMLogin(username, password, token, len);

            if (stat != 1)
            {
                throw new System.Exception("c_SMLogin returned " + stat);
            }
            return token.ToString();
        }
        */
        /// <summary> Returns the cookie SMSESSION </summary>
        /// <returns>string of the value of the SMSESSION cookie</returns>
        public static string GetSMToken()
        {
            if (HttpContext.Current == null) return "";
            if (HttpContext.Current.Request.ServerVariables["HTTP_COOKIE"] == null) return "";
            string cookie = HttpContext.Current.Request.ServerVariables["HTTP_COOKIE"];
            if (cookie == null) return "";
            string smtoken = ParseBetween(cookie, "SMSESSION=", ";");

            if (smtoken == null) smtoken = "";
            if (smtoken.Length <= 600) smtoken = "";
            return smtoken;
        }

        /// <summary> Delete the file with full path that is passed. </summary>
        /// <param name="filename"></param>
        public static void DeleteFile(string filename)
        {
            try
            {
                File.Delete(filename);
            }
            catch (IOException)
            { // Only catch IO Exceptions here. All others can raise normally (i.e. UnAuthorized access)
                System.Threading.Thread.Sleep(3000); // Wait 3 seconds and try again
                try
                {
                    File.Delete(filename); // If this one errors let it raise for now
                }
                catch (Exception e)
                {
                    throw new Exception("File is in use. Please try again later. " + e.Message);
                }
            }
        }

        /// <summary> Add the string to the file specified. Good for adding to the bottom of a log file </summary>
        /// <param name="filename"></param>
        /// <param name="contents"></param>
        public static void AppendFile(string filename, string contents)
        {
            int retries = 0;
            while (true) // pattern to re-try until propery maxRetries
            {
                try
                {
                    using (StreamWriter writer = new StreamWriter(filename, true)) // true = Append
                    {
                        writer.WriteLine(contents);
                        writer.Close();
                    }
                    return; // or break but I read return is faster ;)
                }
                catch (Exception e) // Really should only retry certain types like file locked.
                {	// Yes it is faster to catch the exception on the off chance the dir is not there than calling IfExists() every time
                    if (e.Message.Contains("Could not find a part of the path"))
                    {
                        try
                        {
                            System.IO.Directory.CreateDirectory(Path.GetDirectoryName(filename));
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Error creating directory " + filename + " " + ex.Message);
                        }
                        continue; // No sleep needed so retry now
                    }
                    if (retries++ == maxRetries) throw new Exception(e.Message);
                    System.Threading.Thread.Sleep(retryMilliSeconds);
                }
            }
        }

        /// <summary> Reads a file from disk into a string. If filename does not contain path it tries to read from current web-server directory </summary>
        /// <param name="filename"></param>
        /// <returns>string</returns>
        public static string ReadFile(string filename)
        {
            if (filename.Substring(0, 2) != @"\\" && filename.Substring(1, 1) != ":") // relative or just filename so ask IIS to path it for us!
            {
                if (HttpContext.Current != null) filename = HttpContext.Current.Server.MapPath(filename);
            }
            // This uses a "read" only access so even if file is locked for write you should be able to read it
            try
            {
                return File.ReadAllText(filename);
            }
            catch
            {
                System.Threading.Thread.Sleep(1000);
                return File.ReadAllText(filename);
            }
        }

        /// <summary> Write the string to the file specified. Will over-write if file exisits </summary>
        /// <param name="filename"></param>
        /// <param name="contents"></param>
        public static void WriteFile(string filename, string contents)
        {
            if (!filename.StartsWith(@"\\") && filename.Substring(1, 1) != ":" && HttpContext.Current != null) filename = HttpContext.Current.Server.MapPath(filename); // relative or just filename
            if (filename.Contains("\\")) System.IO.Directory.CreateDirectory(Path.GetDirectoryName(filename));
            using (StreamWriter writer = new StreamWriter(filename, false)) // false = Overwrite
            {
                writer.Write(contents);
                writer.Close();
            }
        }

        /// <summary> Writes to the Event Viewer under the Application folder </summary>
        /// <param name="source"></param>
        /// <param name="contents"></param>
        public void EventLogWrite(string source, string contents)
        {	// Create the source, if it does not already exist.
            if (!EventLog.SourceExists(source))
            {
                EventLog.CreateEventSource(source, "Application"); // Each new source must be "registered"
                System.Threading.Thread.Sleep(4000);
            }
            EventLog myLog = new EventLog();
            myLog.Source = source;
            myLog.WriteEntry(contents, EventLogEntryType.Warning, 3);
        }

        /// <summary> Writes to the Event Viewer under the Application folder </summary>
        /// <param name="source"></param>
        /// <param name="logName"></param>
        public string EventLogRead(string source, string logName = "Application")
        {	// Create the source, if it does not already exist.
            var s = new StringBuilder(1000); // Pre-allocate some bytes

            EventLog eventLog = new EventLog(logName, "."); // . = localhost
            // The EventLog.Source member doesn't work as a filter so use Linq to only get the records matching the source name
            var ev = from EventLogEntry e in (eventLog).Entries where e.Source == source select e; //  && e.Message.Contains("")
            foreach (EventLogEntry entry in ev)
            {
                s.Append(entry.TimeWritten + " " + entry.Message + "\n");
            }
            return s.ToString();
        }

        /// <summary> Writes to the Database in the LOG table </summary>
        /// <param name="source"></param>
        /// <param name="contents"></param>
        public static void LogToDB(string user, string description, string extraData)
        {
            string host = Environment.GetEnvironmentVariable("COMPUTERNAME");
            string type = "UPDATE";
            string desc = description.ToLower();
            if (desc.Contains("delete") || desc.Contains("remove")) type = "DELETE";
            if (desc.Contains("insert")) type = "INSERT";
            try
            {
                host = HttpContext.Current.Request.ServerVariables["REMOTE_HOST"];
                host = System.Net.Dns.GetHostEntry(host).HostName; // Works with IPv4 and 6
            }
            catch { }

            DBHelper db = new DBHelper("select * from Log where 1=2"); // To get the "schema" or just field data
            db.AddRow();
            db.Set("Username", user);
            db.Set("Description", description);
            db.Set("Type", type);
            db.Set("Data", extraData);
            db.Set("Computer", host);
            db.Commit();
        }

        /// <summary> Sends an email using the STMP server set it the web.config or over-ridden with the .STMPServer property </summary>
        /// <param name="toList"></param>
        /// <param name="from"></param>
        /// <param name="subject"></param>
        /// <param name="msg"></param>
        /// <param name="cc"></param>
        public static void SendMail(string toList, string from, string subject, string msg)
        {
            SendMail(toList, from, subject, msg, "");
        }

        /// <summary> Sends an email using the STMP server set it the web.config or over-ridden with the .STMPServer property </summary>
        /// <param name="toList"></param>
        /// <param name="from"></param>
        /// <param name="subject"></param>
        /// <param name="msg"></param>
        /// <param name="cc"></param>
        public static void SendMail(string toList, string from, string subject, string msg, string cc)
        {
            toList = toList.Replace(";", ","); // Yes .net hates semicolons and throws an error
            toList = toList.TrimEnd(',');
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(from, toList, subject, msg);
            message.IsBodyHtml = true;
            cc = cc.Replace(";", ",");
            foreach (string email in cc.Split(','))
            {
                if (email != "") message.CC.Add(new System.Net.Mail.MailAddress(email));
            }
            //message.Priority = System.Net.Mail.MailPriority.High;
            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient(SmtpServer);
            client.Timeout = 3000; // Wow the default is over 1 1/2 minutes for what may be a network hicup
            try
            {
                client.Send(message);
            }
            catch
            {
                System.Threading.Thread.Sleep(1000); // Wait 1 more second and then retry
                try
                {
                    client.Send(message);
                }
                catch
                {
                    //throw new Exception("Error in Util.cs/SendMail(): " + e.Message + " " + e.StackTrace);
                }
            }
        }

        /// <summary> Fetch the contents of the given web-server address </summary>
        /// <param name="url"></param>
        /// <returns>string</returns>
        public static string GetURL(string url, string SMSession = "", string userAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0)") // IE 11
        {
            WebClient wc = new WebClient();
            if (ProxyServer != "" && !url.Contains(".micron.com")) // Do not need to use proxy server for internal addresses
                wc.Proxy = new System.Net.WebProxy(ProxyServer, true);

            string cookie = "SMSESSION=" + SMSession;
            wc.Headers.Add("COOKIE", cookie);

            // Always call this because some proxy server use certs even though the website you are visiting is just http://
            System.Net.ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);

            wc.Headers.Add("User-Agent", userAgent); // Just set something since some servers require this
            return new UTF8Encoding().GetString(wc.DownloadData(url)); // Adjust for UTF8 encoding or regular ASCII
        }

        public static string GetUrlBasicAuth(string url, string username, string password)
        {
            var wc = new WebClient();
            var myCache = new CredentialCache();
            myCache.Add(new Uri(url), "Basic", new NetworkCredential(username, password));
            wc.Credentials = myCache;

            return new UTF8Encoding().GetString(wc.DownloadData(url)); 
        }

        public static string PostURL(string url, string postData, string SMSession = "")
        {
            if (SMSession == "") SMSession = GetSMToken(); // or GetSMSession(username, password);

            System.Net.WebClient wDownload = new System.Net.WebClient();
            string cookie = "SMSESSION=" + SMSession;
            wDownload.Headers.Add("COOKIE", cookie);

            System.Text.UTF8Encoding objUTF8 = new System.Text.UTF8Encoding();
            wDownload.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            byte[] bPostArray = Encoding.UTF8.GetBytes(postData);
            byte[] bHTML = wDownload.UploadData(url, "POST", bPostArray);
            string returnText = objUTF8.GetString(bHTML);

            return returnText;
        }

        public void SaveJPG(System.Drawing.Image image, string szFileName)
        {
            SaveJPG(image, szFileName, 25); // 25 = compression ration of 1 - 100 with 100 = no compression
        }

        public void SaveJPG(System.Drawing.Image image, string szFileName, long lCompression)
        {
            EncoderParameters eps = new EncoderParameters(1);
            eps.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, lCompression);
            ImageCodecInfo ici = GetEncoderInfo("image/jpeg");
            image.Save(szFileName, ici, eps);
        }
        
        /// <summary>Takes a System.Drawing.Image or Bitmap (will implicitly convert) and saves to disk. Uses extension of .jpg, .gif, .bmp, etc. to determine format.</summary>
        /// <example><code>
        /// System.Util obj = new System.Util();
        ///	Bitmap objBitmap = new Bitmap(@"\\webdata\webapps\corp\webservices\fish.jpg");
        /// obj.SaveImage(objBitmap, @"\\webdata\webapps\corp\webservices\fish.gif"); // convert to GIF
        /// </code></example>
        public static void SaveImage(System.Drawing.Image inputImage, string destinationFilename)
        {
            SaveImage(inputImage, destinationFilename, 75);
        }

        /// <summary>Takes a System.Drawing.Image or Bitmap (will implicitly convert) and saves to disk. Uses extension of .jpg, .gif, .bmp, etc. to determine format.</summary>
        /// <example><code>
        /// System.Util obj = new System.Util();
        ///	Bitmap objBitmap = new Bitmap(@"\\webdata\webapps\corp\webservices\fish.gif");
        /// obj.SaveImage(objBitmap, @"\\webdata\webapps\corp\webservices\fish.jpg", 70); // 1-100 with 100 be lossless. Usually 70-80 is good
        /// </code></example>
        public static void SaveImage(System.Drawing.Image inputImage, string destinationFilename, int lCompression)
        {
            Bitmap outputBitMap = new Bitmap(inputImage);//,newPhotoWidth,newPhotoHeight); // or stream
            //Response.Write("01<br>"); Response.Flush();
            EncoderParameters eps = new EncoderParameters(1);
            //Response.Write("02<br>"); Response.Flush();
            eps.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, lCompression);
            //Response.Write("03<br>"); Response.Flush();
            string mime = "image/jpeg";
            //Response.Write("04<br>"); Response.Flush();
            if (System.IO.Path.GetExtension(destinationFilename).ToLower() == ".gif") mime = "image/gif";
            ImageCodecInfo ici = GetEncoderInfo(mime);
            //Response.Write("05<br>"); Response.Flush();
            // Save the newly resized bitmap with the appropriate compression ratio:
            try { outputBitMap.Save(destinationFilename, ici, eps); }
            catch (Exception oErr) { HttpContext.Current.Response.Write(oErr.Message + "<br>"); }
            //Response.Write("06<br>"); Response.Flush();
        }
        

        //Get the codec info for the specified mime type
        private static ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (int j = 0; j < encoders.Length; j++)
            {
                if (encoders[j].MimeType == mimeType) return encoders[j];
            }
            return null;
        }

        //This does both the image resize and the saving with a specified compression ratio:
        public void ResizeImage(System.Drawing.Image inputImage, string destinationFilename)
        {
            int newPhotoWidth = inputImage.Width / 3; // calculate new image size
            int newPhotoHeight = inputImage.Height / 3;
            long lCompression = 20;   //The compression ratio we'll be saving with, from 1 to 100, 100 is the best quality.

            // Create a bitmap with the new size, and with the passed-in image
            Bitmap outputBitMap = new Bitmap(inputImage, newPhotoWidth, newPhotoHeight); // or stream
            EncoderParameters eps = new EncoderParameters(1);
            eps.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, lCompression);
            ImageCodecInfo ici = GetEncoderInfo("image/jpeg");
            // Save the newly resized bitmap with the appropriate compression ratio:
            outputBitMap.Save(destinationFilename, ici, eps);
			outputBitMap.Dispose();
        }

        // This forces the image to a certain size. proportions are maintained with white space 
        public static System.Drawing.Bitmap FixedSize(System.Drawing.Bitmap imgPhoto, int width, int Height)
        {
            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;
            int sourceX = 0;
            int sourceY = 0;
            int destX = 0;
            int destY = 0;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)width / (float)sourceWidth);
            nPercentH = ((float)Height / (float)sourceHeight);
            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
                destX = System.Convert.ToInt16((width - (sourceWidth * nPercent)) / 2);
            }
            else
            {
                nPercent = nPercentW;
                destY = System.Convert.ToInt16((Height - (sourceHeight * nPercent)) / 2);
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap bmPhoto = new Bitmap(width, Height, PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(imgPhoto.HorizontalResolution, imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.White);
            grPhoto.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new Rectangle(destX, destY, destWidth, destHeight),
                new Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            return bmPhoto;
        }
              
        /// <summary>
        /// This writes a URL to a file including binary. This is for getting WAV or JPG files from others servers
        /// </summary>
        /// <param name="url"></param>
        /// <param name="file"></param>
        public static void GetURLToFile(string url, string file)
        {
            byte[] binary = GetURLToByte(url);
            try
            {
                using (FileStream fs = new FileStream(file, FileMode.Create))
                {
                    BinaryWriter w = new BinaryWriter(fs); // Create the writer for data.
                    w.Write(binary);
                    w.Close();
                    fs.Close();
                }
            }catch (Exception e)
            {
                Console.WriteLine("ERROR GetURLToFile: " + e.Message);
            }
        }

        /// <summary>
        /// This downloads a URL into a byte array. Called from GetURLToFile to stream bytes to disk
        /// </summary>
        /// <param name="url"></param>
        /// <returns>byte[]</returns>
        public static byte[] GetURLToByte(string url)
        {
            Uri uri = new Uri(url); // This helps by validating the address and throwing an error if incorrect
            WebClient wDownload = new WebClient();
            wDownload.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; MTI; rv:11.0) like Gecko"); // IE 11 browser

            //CredentialCache myCache = new CredentialCache();
            //wDownload.Credentials = myCache;
            wDownload.UseDefaultCredentials = true;
            try
            {
                return wDownload.DownloadData(uri); // byte[]
            }
            catch (Exception) 
            {
                System.Threading.Thread.Sleep(3000); // Wait 3 seconds and try again
                return wDownload.DownloadData(uri); // byte[]
            }
        }

        /// <summary>
        /// Returns true if string matches the passed regular expression pattern case insensitive
        /// </summary>
        /// <param name="str"></param>
        /// <param name="pattern"></param>
        /// <returns>bool</returns>
        public static bool IsMatchRegExp(string str, string pattern)
        {
            Util myUtil = new Util();
            return myUtil.IsMatchRegExp(str, pattern, IgnoreCase: true);
        }

        /// <summary>
        /// Returns true if string matches the passed regular expression pattern
        /// </summary>
        /// <param name="str"></param>
        /// <param name="pattern"></param>
        /// <param name="IgnoreCase"></param>
        /// <returns>bool</returns>
        private bool IsMatchRegExp(string str, string pattern, bool IgnoreCase)
        {
            System.Text.RegularExpressions.RegexOptions flag;
            System.Text.RegularExpressions.Regex emailregex;
            if (IgnoreCase)
            {
                flag = System.Text.RegularExpressions.RegexOptions.IgnoreCase;
                emailregex = new System.Text.RegularExpressions.Regex(pattern, flag);
            }
            else
            {
                emailregex = new System.Text.RegularExpressions.Regex(pattern);
            }
            return emailregex.IsMatch(str);
        }

        

        /// <summary>
        /// Returns true if string can be converted to a number
        /// </summary>
        /// <param name="str"></param>
        /// <returns>bool</returns>
        public static bool IsNumeric(object str)
        {
            return IsDecimal(str);
        }

        public static bool IsDecimal(object str)
        {
            decimal result;
            return Decimal.TryParse(str.ToString(), out result);
        }

        /// <summary> Return true if whatever is passed can be converted to an integer </summary>
        /// <param name="stringToTest"></param>
        /// <returns></returns>
        public static bool IsInteger(object stringToTest)
        {
            int result;
            return int.TryParse(stringToTest.ToString(), out result);
        }
        public static bool IsInt(object stringToTest)
        {
            return IsInteger(stringToTest);
        }

        /// <summary>
        /// Make sure string passed has a @ and a dot at least (fred@micron = false, fredat.com = false, fred@.com = false)
        /// </summary>
        /// <param name="email"></param>
        /// <returns>bool</returns>
        public static bool IsEmailAddress(string email)
        {
            return IsMatchRegExp(email, "([^@]+)@((.+)\\.(.+))");
        }

        /// <summary>
        /// Convert the object passed to a DateTime and if successful return true
        /// </summary>
        /// <param name="inValue"></param>
        /// <returns></returns>
        public static bool IsDate(object inValue)
        {
            try
            {
                System.DateTime.Parse(inValue.ToString());
            }
            catch
            {
                return false;
            }
            return true; // Else it was successfully converted to a date
        }

        /// <summary>Returns true if a valid Microsoft GUID is passed. Can have {} or with or without dashes (36 chars with the dashes) </summary>
        /// <param name="inValue"></param>
        /// <returns>bool</returns>
        public static bool IsGUID(object inValue)
        {
            if (inValue == null) return false;
            string s = inValue.ToString();
            try
            {
                new Guid(s); // Cannot use Base64 encoded like PyUE4E+JEdOaDAMF6CwzAQ==. Decode first with Convert.FromBase64String(string);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }


        /// <summary>Returns true if all characters are uppercase</summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static bool IsAllUpper(string input)
        {
            for (int i = 0; i < input.Length; i++)
                if (Char.IsLetter(input[i]) && !Char.IsUpper(input[i])) return false;
            return true;
        }

        /// <summary>
        /// Returns true if the passed object can be consided a boolean including 1 or "1" or "True"
        /// </summary>
        /// <param name="o"></param>
        /// <returns>bool</returns>
        public static bool ToBool(object o)
        {
            if (o == null) return false;
            string s = o.ToString();
            if (s == "") return false;
            if (s == "0") return false;
            if (s == "1") return true;
            return Convert.ToBoolean(s); // .NET needs the word "True" or "False" thus the checking above for 0 and 1
        }

        /// <summary>
        /// Returns true if the passed object can be consided a boolean including 1 or "1" or "True"
        /// </summary>
        /// <param name="o"></param>
        /// <returns>bool</returns>
        public static DateTime ToDate(object o)
        {
            if (o == null) return DateTime.MinValue;
            string s = o.ToString();
            DateTime retDate = DateTime.Parse(s);
            return retDate;
        }


        /// <summary>
        /// Returns integer value if the passed object can be converted to an integer. Null and blank return 0
        /// </summary>
        /// <param name="str"></param>
        /// <returns>int</returns>
        public static int ToInt(object str)
        {
            if (str == null || str.ToString() == "") return 0;
            if (str is double) return Convert.ToInt32(str);
            return Convert.ToInt32(str.ToString());
        }
        public static double ToDouble(object str)
        {
            if (str == null || str.ToString() == "") return 0;
            return Convert.ToDouble(str.ToString());
        }
        public static string ToDecimal(object str)
        {
            if (str.ToString() == "") return "";
            return Convert.ToDouble(str).ToString("G0");
        }

        /// <summary>
        /// Return the contents of a string between the start and end string case insensitive
        /// </summary>
        /// <param name="inString"></param>
        /// <param name="startText"></param>
        /// <param name="endText"></param>
        /// <returns>string</returns>
        public static string ParseBetween(string inString, string startText, string endText)
        {
            return LeftOf(RightOf(inString, startText), endText);
            // If this was not case-INsenstive we could use this:
            /*
            var list = inString.Split(new string[] { startText, endText }, StringSplitOptions.None);
            if (list.Length >= 1) return list[1].Trim();
            return "";
            */
        }

        /// <summary>This splits on both start and end and returns the middle element. Ensure end text does does come before the start text</summary>
        /// <param name="inString"></param>
        /// <param name="startText"></param>
        /// <param name="endText"></param>
        /// <returns>string</returns>
        public static string Between(string inString, string startText, string endText)
        {
            var list = inString.Trim().Split(new string[] { startText, endText }, StringSplitOptions.None); // NOT CASE-SENSITIVE!
            if (list.Length > 1) return list[1].Trim();
            return "";
        }

        /// <summary>
        /// Returns true if string1 contains string2 case-insensitive
        /// </summary>
        /// <param name="bigString"></param>
        /// <param name="searchfor"></param>
        /// <returns>boolean</returns>
        public static bool Contains(string bigString, string searchString)
        {
            return bigString.IndexOf(searchString, ignoreCase) >= 0;
        }

        /// <summary>
        /// Return all characters to the right of a string. Possible case insensitive
        /// </summary>
        /// <param name="str"></param>
        /// <param name="searchfor"></param>
        /// <returns>string</returns>
        public static string RightOf(string str, string searchfor, bool wholeStringIfNotFound = false)
        {
            StringComparison compare = StringComparison.CurrentCultureIgnoreCase;
            if (CaseSensitive) compare = StringComparison.CurrentCulture;
            int pos = str.IndexOf(searchfor, compare);
            if (pos >= 0) return str.Substring(pos + searchfor.Length).Trim();
            if (wholeStringIfNotFound)
                return str.Trim();
            else
                return "";
        }

        /// <summary>
        /// Return the portion of the 1st string left of the 2nd string
        /// </summary>
        /// <param name="str"></param>
        /// <param name="chr"></param>
        /// <returns>string</returns>
        public static string LeftOf(object stringOrObject, string chr, bool wholeStringIfNotFound = false)
        {
            string str = stringOrObject.ToString();
            StringComparison compare = StringComparison.CurrentCultureIgnoreCase;
            if (CaseSensitive) compare = StringComparison.CurrentCulture;

            int leftOfPos = str.IndexOf(chr, compare);
            if (leftOfPos != -1) return Left(str, leftOfPos).Trim();
            if (wholeStringIfNotFound)
                return str.Trim();
            else
                return "";
        }

        /// <summary>
        /// Returns left most characters but you do not need to check the length and works with any object (that can be converted to a string)
        /// </summary>
        /// <param name="str"></param>
        /// <param name="leftOfPos"></param>
        /// <returns>string</returns>
        public static string Left(object searchStr, int leftOfPos)
        {
            string str = searchStr.ToString();
            if (leftOfPos >= str.Length) return str;
            return str.Substring(0, leftOfPos);
        }

        /// <summary>Like String.Replace() but case INsensitive! </summary>
        /// <param name="str"></param>
        /// <param name="pattern"></param>
        /// <param name="replaceWith"></param>
        /// <param name="IgnoreCase">true</param>
        /// <returns>string</returns>
        public string Replace(string str, string pattern, string replaceWith, bool IgnoreCase = true)
        {
            if (String.IsNullOrEmpty(str)) return "";
            pattern = pattern.Replace("*", "\\*"); // Just incase they want to replace *NAME*
            System.Text.RegularExpressions.RegexOptions flag = System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace;
            if (IgnoreCase) flag = System.Text.RegularExpressions.RegexOptions.IgnoreCase;
            return System.Text.RegularExpressions.Regex.Replace(str, pattern, replaceWith, flag);
        }

        /// <summary>
        /// Takes a string and converts to XML with error checking and gets the value of the node/element passed (1st if multiple)
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="nodeName"></param>
        /// <returns>string </returns>
        public static string GetNode(string xml, string nodeName)
        {
            XmlDocument doc = new XmlDocument();
            try {
                doc.LoadXml(xml);
            } 
            catch (Exception e) {
                throw (new Exception("Bad XML: " + e.Message + "\n<xmp>" + xml + "</xmp>")); // Yes XMP tags will make it actually display in a browser. Kinda like <!DATA
            }

            var node = doc.SelectSingleNode("//" + nodeName);  // The // means find anywhere
            if (node == null) return nodeName + " not found";
            return node.InnerXml; // Get just the contents (between the start and end tags)
        }

        /// <summary>Given a XmlDocument instance look for the first occurance of the passed element</summary>
        /// <param name="xml"></param>
        /// <param name="nodeName"></param>
        /// <returns></returns>
        public static string GetNode(XmlDocument doc, string XPATH)
        {
            if (doc == null || doc.FirstChild.ChildNodes.Count == 0) throw new Exception("XML is empty");

            var node = doc.SelectSingleNode(XPATH);  // The // means find anywhere
            if (node == null) return XPATH + " not found";
            return node.InnerXml; // Get just the contents (between the start and end tags)
        }

        public static int GetNodeInt(XmlDocument doc, string XPATH)
        {
            if (doc == null || doc.FirstChild.ChildNodes.Count == 0) throw new Exception("XML is empty");

            var node = doc.SelectSingleNode(XPATH);  // The // means find anywhere
            if (node == null) return 0;
            return Convert.ToInt32(node.InnerText);
        }

        public static string GetAttr(string xml, string attribute)
        {
            return ParseBetween(xml, attribute + "=\"", "\"");
        }

        /// <summary> Returns the value for the given Environment variable like COMPUTERNAME </summary>
        /// <param name="variable"></param>
        /// <returns></returns>
        public static string GetEnv(object environmentVariable)
        {
            string env = environmentVariable.ToString();
            string results = ""; // Do we return blank or throw an error if the env variable passed is not found?
            try {
                results = System.Environment.GetEnvironmentVariable(env);
            }
            catch { } // Just return blank if not found
            return results;
        }

        /// <summary>Given a directory returns an array of strings objects for each filename in the directory</summary>
        /// <param name="directory"></param>
        /// <param name="filter"></param>
        /// <returns>Collection of strings (IEnumerable of string)</returns>
        public static IEnumerable<string> GetFiles(string directory, string filter = "*.*", bool recurseSubDirs = true)
        {
            var searchOption = SearchOption.TopDirectoryOnly;
            if (recurseSubDirs) searchOption = SearchOption.AllDirectories;
            //return Directory.EnumerateFiles(directory, filter, searchOption);

            var di = new DirectoryInfo(directory);
            var files = (from file in di.EnumerateFiles(filter, searchOption)
                        orderby file.CreationTime ascending
                        select file.FullName).Distinct(); // Don't need <string> here, since it's implied
            return files;
        }

        /// <summary> Uses a JavaScript pop-up to display messages </summary>
        /// <param name="errorMsg"></param>
        public static void AlertAndEnd(string errorMsg, string redirectURL="")
        {
            Print("</select></form></table></table>"); // just in-case they are in the middle of a tag
            Print("<meta http-equiv='Pragma' content='no-cache'><meta http-equiv='Expires' content='0'>");
            errorMsg = errorMsg.Replace(@"\", @"\\");  // Escape back-slashes for JavsScript alert
            errorMsg = errorMsg.Replace("\n", "\\n"); // Newlines needs to be "\n" for JavaScript
            errorMsg = errorMsg.Replace("\r", ""); // CR needs to go away for JavaScript
            errorMsg = errorMsg.Replace("&L10", "");  // Char 10 needs to go away
            errorMsg = errorMsg.Replace("\t", "   "); // Change any tabs to 3 spaces
            errorMsg = errorMsg.Replace("'", "`");    // Change any single quotes for Javascript inclusion
            Print("<script>alert('" + errorMsg + "');</scr" + "ipt>");
            if (redirectURL == "")
                Print("<script>history.go(-1); </scr" + "ipt>");
            else
                Print("<script>location.href='" + redirectURL + "';</scr" + "ipt>");

            Print("</body></html>");
            HttpContext.Current.Response.End();
        }

        /// <summary> Shortcut for printing to the browser (Response.Write) from code-behind. Includes a newline at end </summary>
        /// <param name="str"></param>
        public static void Print(object str)
        {
            System.Web.HttpResponse Response = HttpContext.Current.Response;
            Response.Write(str.ToString() + "\n");
        }

        /// <summary> Return the value (trimed) of a URL parameter like ?name=ZZZZZZ. If not there just returns "" </summary>
        /// <param name="key"></param>
        /// <returns>string</returns>
        public string GetRequest(string key)
        {
            return Request(key);
        }

        public static string EncodeRequest (string key)
        {
            return Request(key).Replace("'", "");
        }        
        
        /// <summary>Returns the data for the specific key passed</summary>
        /// <param name="key">form field name or url parameter</param>
        /// <returns>string value of the requested Form or URL name passed</returns>
        public static string Request(string key)
        {
            if (HttpContext.Current == null) return "";

            if (HttpContext.Current.Request[key] == null) return "";
            if (HttpContext.Current.Request.QueryString[key] != null) // Look on URL first before larger .Form collection and .ServerVariables collection
                return HttpContext.Current.Request.QueryString[key].ToString().Trim();

            return HttpContext.Current.Request[key].ToString().Trim();
        }

        /// <summary> Return the value (trimed) of a Session variable. If not there just returns "" </summary>
        /// <param name="key"></param>
        /// <returns>string</returns>
        public static string GetSession(object key)
        {
            if (HttpContext.Current == null) return "";

            if (HttpContext.Current.Session[key.ToString()] == null)
                return "";
            return HttpContext.Current.Session[key.ToString()].ToString().Trim();
        }

        /// <summary>Returns the value of a cookie and blank if it does not exist </summary>
        ///	<param name="name">Cookie name</param>
        /// <returns> string with the value</returns>
        public static string GetCookie(string name)
        {
            if (HttpContext.Current == null) return "";

            if (HttpContext.Current.Request.Cookies[name] == null) return "";
            return HttpContext.Current.Request.Cookies[name].Value;
        }

        public static void SetCookie(string name, string value, int daysToExpire=0)
        {
            HttpCookie objCookie = new HttpCookie(name);
            objCookie.Value = value;
            if (daysToExpire == 0)
            {
                System.DateTime dt = System.DateTime.Now;
                TimeSpan ts = new TimeSpan(0, 0, 60, 0);
                objCookie.Expires = dt.Add(ts);
            }
            else
            {
                System.DateTime dtExpiry = System.DateTime.Now.AddDays(daysToExpire);
                objCookie.Expires = dtExpiry;
            }

            HttpContext.Current.Response.Cookies.Add(objCookie);
        }
        
        /// <summary>Returns true or false if the user is a member of the Exchange group passed</summary>
        ///	<param name="group">Name or MtGroup or Exchange group</param>
        /// <returns> string with the value</returns>
        public static bool IsGroupMember(string group, string username = "")
        {
            if (username == "") username = GetUsername();

            var sql = @"SELECT Username FROM Groups INNER JOIN GroupMembers ON GroupMembers.GroupNo = Groups.GroupNo
                WHERE GroupName = '" + group + @"'
                OR ParentGroupNo = (SELECT GroupNo FROM Groups WHERE GroupName='" + group + "')";
            var db = new DBHelper(sql, ConfigurationManager.AppSettings["PT"]);
            if (db.Count == 0) return false;
            var myRowArray = db.DataTable.Select("Username='" + username + "'");
            if (myRowArray.Length > 0) return true;
            return false;
        }

        /// <summary>
        /// This returns usernames from OUR system. i.e. The Groups table for DesignRules
        /// </summary>
        /// <param name="groups"></param>
        /// <returns></returns>
        public static List<string> GroupMembers(string groups)
        {
            List<string> allMembers = new List<string>();

            // 1st look in our database for the group. If not found THEN go to Active Directory
            groups = groups.Replace(",", "','"); // Convert for SQL In Clause
            /*
            var sql = @"SELECT DISTINCT Username FROM Groups INNER JOIN GroupMembers ON GroupMembers.GroupNo = Groups.GroupNo 
                WHERE GroupName IN ('" + groups + @"')
                OR ParentGroupNo = (SELECT GroupNo FROM Groups WHERE GroupName='" + groups + @"')
                OR ParentGroupNo IN (SELECT GroupNo FROM Groups WHERE ParentGroupNo=(SELECT GroupNo FROM Groups WHERE GroupName='" + groups + "'))";
            */

            var sql = @"WITH CTE as
            (   -- Anchor query 
	            SELECT Username, Groups.GroupNo FROM Groups
	            LEFT OUTER JOIN GroupMembers On GroupMembers.GroupNo = Groups.GroupNo
	            WHERE GroupName IN ('" + groups + @"')
  
              UNION ALL --bind the Anchor and Recursive queries together.
  
	            SELECT GroupMembers.Username, Groups.GroupNo FROM Groups
	            INNER JOIN GroupMembers On GroupMembers.GroupNo = Groups.GroupNo
	            INNER JOIN CTE ON CTE.GroupNo =  Groups.ParentGroupNo -- The outer or anchor query
             )
            SELECT DISTINCT Username From CTE WHERE Username IS NOT NULL ORDER BY Username";

            var db = new DBHelper(sql);
            if (db.Count > 0) // Then we have it in OUR database
            {
                foreach (System.Data.DataRow myRow in db.Rows) allMembers.Add(myRow["username"].ToString());
            }

            return allMembers;
        }

        /// <summary>Return a single Attribute value from Active Directory given the property name</summary>
        /// <param name="username"></param>
        /// <param name="property"></param>
        /// <returns></returns>
        /*
        public static string UserProperties(string username, string property)
        {
            var ctx = new PrincipalContext(ContextType.Domain);
            var userPrincipal = UserPrincipal.FindByIdentity(ctx, username);
            var directoryEntry = userPrincipal.GetUnderlyingObject() as System.DirectoryServices.DirectoryEntry;

            // Look for the specific Property passed
            if (directoryEntry.Properties.Contains(property))
                return directoryEntry.Properties[property].Value.ToString();
            else
                return String.Empty;
        }
        */
        /// <summary>Given a MTGroup or Exchange group returns a List of usernames</summary>
        ///	<param name="group">name</param>
        /// <returns>List of usernames</returns>
        /*
        public static List<string> GroupDirectMembers(string group)
        {
            // See: http://stackoverflow.com/questions/5895128/attempted-to-access-an-unloaded-appdomain-when-using-system-directoryservices
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "WINNTDOM");

            GroupPrincipal findGroup = GroupPrincipal.FindByIdentity(ctx, group);
            if (findGroup == null) throw new Exception("Group name " + group + " not found.");
            List<string> members = new List<string>();
            foreach (Principal p in findGroup.GetMembers())
            {
                members.Add(p.Name); 
                // var theUser = p as UserPrincipal; // This class has .Active if needed
            }
            return members;
        }

        /// <summary>
        /// Given an MtGroup/Active Directory group returns all usernames
        /// </summary>
        /// <param name="group"></param>
        /// <returns></returns>
        public static List<string> GroupEventualMembers(string group)
        {
            //PrincipalContext ctx = new PrincipalContext(ContextType.Domain); // Bug. See: http://stackoverflow.com/questions/5895128/attempted-to-access-an-unloaded-appdomain-when-using-system-directoryservices
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "WINNTDOM");

            GroupPrincipal findGroup = GroupPrincipal.FindByIdentity(ctx, group);
            if (findGroup == null) throw new Exception("Group name " + group + " not found.");
            List<string> members = new List<string>();
            foreach (var p in findGroup.GetMembers())
            {
                if (p is GroupPrincipal)
                {
                    var groupName = p.Name;
                    foreach (string userName in GroupEventualMembers(groupName))
                    {
                        if (!members.Contains(userName)) members.Add(userName);
                    }
                }
                else
                {
                    members.Add((p as UserPrincipal).Name);
                }
            }
            return members;
        }

        public static List<int> GroupWorkerNos(string group)
        {
            //PrincipalContext ctx = new PrincipalContext(ContextType.Domain); // Bug. See: http://stackoverflow.com/questions/5895128/attempted-to-access-an-unloaded-appdomain-when-using-system-directoryservices
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "WINNTDOM");

            GroupPrincipal findGroup = GroupPrincipal.FindByIdentity(ctx, group);
            if (findGroup == null) throw new Exception("Group name " + group + " not found.");
            List<int> members = new List<int> ();
            foreach (var p in findGroup.GetMembers())
            {
                if (p is GroupPrincipal)
                {
                    var groupName = p.Name;
                    foreach (int workerNo in GroupWorkerNos(groupName))
                    {
                        if (!members.Contains(workerNo)) members.Add(workerNo);
                    }
                }
                else
                {
                    members.Add(Convert.ToInt32((p as UserPrincipal).EmployeeId));
                }
            }
            return members;
        }


        public static bool ReportsTo(string manager, string username)
        {
            var db = new DBHelper("SELECT rtrim(micron_username) as Username FROM SAP_worker WHERE micron_username != '' AND " +
              "eventual_rpt_to_code LIKE (SELECT eventual_rpt_to_code FROM SAP_worker WHERE micron_username='" + manager + "') + '%' " +
              "AND micron_username = '" + username + "'");

            if (db.Count == 0) return false;
            return true;
        }

        /// <summary>Given a MTGroup or Exchange group returns true if it exists</summary>
        ///	<param name="group">name</param>
        /// <returns>True or False</returns>
        public static bool IsGroupValid(string group)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain); // It knows to use the same DC as logged into
            GroupPrincipal findGroup = GroupPrincipal.FindByIdentity(ctx, group);
            if (findGroup == null)
                return false;
            else
                return true;
        }
*/
        /// <summary>Takes a command like cmd /C dir C: or ping localhost -n 1 and return the output from standard out as a string</summary>
        /// <example><code>
        /// var results = Util.Shell("cmd /C dir C:");
        /// </code></example>
        public static string Shell(string command, string currentDirectory = @"C:\Windows\system32")
        {
            string args = "";
            if (command.IndexOf(" ") != -1)
            {
                if (command.StartsWith("\""))
                {
                    args = command.Substring(command.IndexOf("\"", 2) + 2);
                    command = command.Substring(1, command.IndexOf("\"", 2) - 1);
                } else {
                    args = command.Substring(command.IndexOf(" ", 2) + 1);
                    command = command.Substring(0, command.IndexOf(" ", 2));
                }

            }
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.WorkingDirectory = currentDirectory;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.FileName = command;
            p.StartInfo.Arguments = args;
            p.Start();
            return p.StandardOutput.ReadToEnd(); // waits for standard out to finish
        }

        public static void ShellWait(string command, int sleepInMS=1000)
        {
            string args = "";
            if (command.IndexOf(" ") != -1)
            {
                if (command.StartsWith("\""))
                {
                    args = command.Substring(command.IndexOf("\"", 2) + 2);
                    command = command.Substring(1, command.IndexOf("\"", 2) - 1);
                }
                else
                {
                    args = command.Substring(command.IndexOf(" ", 2) + 1);
                    command = command.Substring(0, command.IndexOf(" ", 2));
                }

            }
            //Process p = new Process(); // System.Diagnositics
            using (Process p = new Process())
            {
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.WorkingDirectory = @"C:\Windows\system32";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.FileName = command;
                p.StartInfo.Arguments = args;
                p.Start();
                //p.WaitForExit();
                p.Close();
            }

            System.Threading.Thread.Sleep(sleepInMS);
        }


        /// <summary>Given a file path, Returns the first line that matches (contains the text passed)</summary>
        public static string FindInFile(string filePath, string lookFor)
        {
            foreach (var str in File.ReadLines(filePath).Where(s => s.Contains(lookFor)))
                return str; // Stops at 1st occurance. Where is lazy initialized for performance
            return "";
        }

        /// <summary>Return the first line that matches (assuming newline separated file passed as a string)</summary>
        public static string FindLine(string str, string lookFor)
        {
            foreach (string line in str.Split('\n'))
            {
                if (line.Contains(lookFor)) return line;
            }
            return "";
        }

        /// <summary>Return all lines that match (assuming newline separated file passed as a string)</summary>
        public static List<string> FindLines(string str, string lookFor)
        {
            List<string> returnList = new List<string>();
            foreach (string line in str.Split('\n'))
            {
                if (line.Contains(lookFor)) returnList.Add(line);
            }
            return returnList;
        }

        /// <summary> Formats any XML with newlines and indents (Tidy Format) suitable for displaying </summary>
        /// <param name="xml"></param>
        /// <returns>string</returns>
        public static string FormatXML(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);

            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            XmlTextWriter writer = new XmlTextWriter(sw);
            writer.Formatting = Formatting.Indented;
            doc.WriteTo(writer);
            writer.Close();
            return sb.ToString();
        }

        /// <summary>Replace the inner text of first occurance of the node passed</summary>
        /// <param name="xml">valid XML as a string</param>
        /// <param name="nodeSelector">can be any XPATH i.e. /root/data/groupname</param>
        /// <param name="newValue"></param>
        /// <returns>changed XML</returns>
        public static string XMLSingleReplacer(string xml, string nodeSelector, string newValue)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml); // we should get rid of any xmlns="" (namespace attributes)
            XmlNode myNode = doc.SelectSingleNode(nodeSelector);
            if (myNode != null) myNode.InnerText = newValue;
            return doc.InnerXml;
        }

        // nodeSelector can be any XPATH i.e. /root/data/groupname
        /// <summary>Replace all found occurances of the node passed or the attribute passed (optional)</summary>
        /// <param name="xml">valid XML as a string</param>
        /// <param name="nodeSelector">can be any XPATH i.e. /root/data/groupname</param>
        /// <param name="newValue"></param>
        /// <param name="attribute"></param>
        /// <returns></returns>
        public static string XMLReplacer(string xml, string nodeSelector, string newValue, string attribute = "")
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNodeList aNodes = doc.SelectNodes(nodeSelector);

            foreach (XmlNode aNode in aNodes) // loop through all found nodes
            {
                if (attribute == "") // They want the element "contents" changes not an attribute of it
                    aNode.InnerText = newValue;
                else // find the attribute name they passed us
                {
                    XmlAttribute myAttribute = aNode.Attributes[attribute];
                    if (myAttribute != null) myAttribute.Value = newValue;
                }
            }
            return doc.InnerXml;
        }

        /// <summary>Add passed number of work-days to passed date and return next business day </summary>
        /// <param name="startDate"></param>
        /// <param name="workDays"></param>
        /// <returns>New date as a string</returns>
        public static string AddWorkDays(DateTime startDate, int workDays)
        {   // Since the parameter may be negative we need to set a conversion variable to -1 if it is
            int negativeDays = 1;
            if (workDays < 0)
            {
                negativeDays = -1;
                workDays = Math.Abs(workDays); // Absolute Value
            }
            for (int i = 1; i <= workDays; i++)
            {
                startDate = startDate.AddDays(1 * negativeDays); // move the day forward one
                DayOfWeek weekDay = startDate.DayOfWeek; // So we can check for weekends
                if (negativeDays == -1 && weekDay == DayOfWeek.Sunday)
                    startDate = startDate.AddDays(-2);

                if (negativeDays == 1 && weekDay == DayOfWeek.Saturday)
                    startDate = startDate.AddDays(2);
            }
            // Since we do not want to have an end date on a saturday or sunday call routine that checks and bumps it to Monday
            return NextWorkingDay(startDate).ToString("M/d/yyyy");
        }

        /// <summary>If date passed is a Sat. or Sun. return the next monday else return same date </summary>
        /// <param name="startDate"></param>
        /// <returns>New Date</returns>
        public static DateTime NextWorkingDay(DateTime startDate)
        {
            DayOfWeek weekDay = startDate.DayOfWeek;
            if (weekDay == DayOfWeek.Saturday) startDate = startDate.AddDays(2);
            if (weekDay == DayOfWeek.Sunday) startDate = startDate.AddDays(1);
            return startDate;
        }

        /// <summary>Returns the number of work days for the # of calendar days passed. i.e Return 5 for 7 or 10 for 14</summary>
        /// <param name="days"></param>
        /// <returns>number of days</returns>
        public static int DaysToWorkDays(int days)
        {   // If the days are greater than 5 we have a much bigger chance of going over a weekend
            if (days <= 5) { return days; } // 5 or less should not change since we do not know the start date
            decimal workDays = days * 5 / 7;
            return Convert.ToInt32(Math.Ceiling(workDays)); // always round up to the nearest full day
        }

        public static decimal Round(object s, int decimalPlaces)
        {
          if (s == null || s.ToString() == "") return 0;
          try {
            return Math.Round(Convert.ToDecimal(s.ToString()), decimalPlaces);
          } catch  {
            return 0;
          }
        }
        
        public static string FormatNumber(Object o)
        {
            double MyDouble = Convert.ToDouble(o);
            return MyDouble.ToString("n0");
        }        

        public static string FormatCurrency(Object o)
        {
            double MyDouble = Convert.ToDouble(o);
            return String.Format("{0:c}", MyDouble);
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

        // This would need System.Data.DataSetExtensions referenced
        /*
        private T JsonDeserialize<T>(string jsonString)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
            MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
            return (T)serializer.ReadObject(ms);
            //return obj;
        }
        */

        /// <summary>Given an object that can be serialized, returns the XML as a string </summary>
        /// <param name="obj">Any object that support ISerialable interface</param>
        public static string SerializeObject(object obj)
        {
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(obj.GetType());
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                serializer.Serialize(ms, obj);
                ms.Position = 0;
                xmlDoc.Load(ms);
                return xmlDoc.InnerXml;
            }
        }

        /// <summary>Create a CQ record by sending a http request to their web-service</summary>
        /// <param name="CQDepartmentName"></param>
        /// <param name="CQSubGroup"></param>
        /// <param name="micronSystem"></param>
        /// <param name="title"></param>
        /// <param name="desc"></param>
        /// <param name="contact"></param>
        /// <param name="CaseType"></param>
        /// <param name="severity"></param>
        /// <param name="notification"></param>
        /// <returns></returns>
        public static string CreateClearQuest(string CQDepartmentName, string CQSubGroup, string system, string title, string desc, string contact = "", string severity="3-Normal")
        {
            string url = "http://cqwebprod.micron.com/webapps/is/cpat/smg/cqreports/cqwebservices/cqcreatechangerequest.asmx/CreateChangeRequest";
            if (contact == "") contact = Util.GetUsername();
            string formData = "DataBase=ISTRK&Schema=ISTRK&ChangeRequestType=WorkRequest&Username=" + contact;
            
            // Example URL (GET Request
            // "&FieldNames=Headline&FieldNames=Description&FieldNames=Contact&FieldNames=ContactEmpNo&FieldNames=CustomerSeverity&FieldNames=DepartmentDBID&FieldNames=SubGroupDBID&FieldNames=MicronSystemDBID&FieldValues=Testing+Headline&FieldValues=Testing+API+for_Clearquest&FieldValues=rchipman&FieldValues=6356&FieldValues=3-Normal&FieldValues=ASSM&FieldValues=ASSM+Documentation&FieldValues=ASM+PDSS";

            formData += "&FieldNames=Headline&FieldValues=" + HttpUtility.UrlEncode(title);
            formData += "&FieldNames=Description&FieldValues=" + HttpUtility.UrlEncode(desc);
            formData += "&FieldNames=Contact&FieldValues=" + contact;
            formData += "&FieldNames=ContactEmpNo&FieldValues=" + Util.GetEmpNoFromUsername(contact);
            formData += "&FieldNames=CustomerSeverity&FieldValues=" + severity;
            formData += "&FieldNames=DepartmentDBID&FieldValues=" + HttpUtility.UrlEncode(CQDepartmentName);
            formData += "&FieldNames=SubGroupDBID&FieldValues=" + HttpUtility.UrlEncode(CQSubGroup);
            formData += "&FieldNames=MicronSystemDBID&FieldValues=" + HttpUtility.UrlEncode(system);
            //formData += "&FieldNames=Contact_CCList&FieldValues=";
            //formData += "&FieldNames=SendNotification&FieldValues=YES"; // Default anyway

            string s = PostURL(url, formData);
            string istrk = GetXmlValue(s, "Value");
            string errMsg = GetXmlValue(s, "Error");
            if (errMsg == "errors")
            {
                s = s.Substring(s.IndexOf("</Error>") + 8);
                errMsg = GetXmlValue(s, "Error");
            }

            return istrk + " " + errMsg;
        }

        /// <summary>Parses XML (or HTML) for a single element (or tag) anywhere in the string</summary>
        /// <param name="xml"></param>
        /// <param name="nodeName"></param>
        /// <returns>Value that was between the start and end tag</returns>
        public static string GetXmlValue(string xml, string nodeName) // Should this have an option to be case-insensitive?
        {
            int start = xml.IndexOf("<" + nodeName + ">") + nodeName.Length + 2;
            int endof = xml.IndexOf("</" + nodeName + ">");
            if (endof > start)
                return xml.Substring(start, endof - start);
            else
                return "";
        }

        /// <summary>Return SELECTED if 2 values mach. Util.IsSelected(matID, id) </summary>
        public static string IsSelected(string str1, string str2)
        {
            if (str1.Trim().ToUpper() == str2.Trim().ToUpper())
                return " SELECTED";
            else
                return "";
        }

        /// <summary>Given a single XmlNode instance return the full ancenstral path</summary>
        /// <param name="node"></param>
        /// <returns></returns>
        static string GetXpath(XmlNode node)
        {
            if (node.Name == "#document") return "";
            return GetXpath(node.SelectSingleNode("..")) + "/" + (node.NodeType == XmlNodeType.Attribute ? "@" : "") + node.Name;
        }
        
        /// <summary></summary>
        /// <param name="XmlDocument"></param>
        /// <param name="XPATH"></param>
        /// <returns>integer</returns>
        int NodesAsInt(XmlDocument xdoc, string XPATH)
        {
            int qty = 0;
            var nodes = xdoc.SelectNodes(XPATH);
            foreach (XmlNode theNode in nodes) qty += Convert.ToInt32(theNode.InnerText);
            return qty;
        }
    }
}

// Can be called from .Sort() or anything that takes a System.Collections.IComparer
public class AlphanumComparatorFast : System.Collections.IComparer
{
    public int Compare(object x, object y)
    {
        string s1 = x as string;
        if (s1 == null) return 0;
        string s2 = y as string;
        if (s2 == null) return 0;

        int len1 = s1.Length;
        int len2 = s2.Length;
        int marker1 = 0;
        int marker2 = 0;

        // Walk through two the strings with two markers.
        while (marker1 < len1 && marker2 < len2)
        {
            char ch1 = s1[marker1];
            char ch2 = s2[marker2];

            // Some buffers we can build up characters in for each chunk.
            char[] space1 = new char[len1];
            int loc1 = 0;
            char[] space2 = new char[len2];
            int loc2 = 0;

            // Walk through all following characters that are digits or
            // characters in BOTH strings starting at the appropriate marker.
            do
            {
                space1[loc1++] = ch1;
                marker1++;

                if (marker1 < len1)
                    ch1 = s1[marker1];
                else
                    break;
            } while (char.IsDigit(ch1) == char.IsDigit(space1[0]));
            do
            {
                space2[loc2++] = ch2;
                marker2++;

                if (marker2 < len2)
                    ch2 = s2[marker2];
                else
                    break;
            } while (char.IsDigit(ch2) == char.IsDigit(space2[0]));

            // If we have collected numbers, compare them numerically.
            // Otherwise, if we have strings, compare them alphabetically.
            string str1 = new string(space1);
            string str2 = new string(space2);

            int result;

            if (char.IsDigit(space1[0]) && char.IsDigit(space2[0]))
            {
                int thisNumericChunk = int.Parse(str1);
                int thatNumericChunk = int.Parse(str2);
                result = thisNumericChunk.CompareTo(thatNumericChunk);
            }
            else
                result = str1.CompareTo(str2);

            if (result != 0) return result;
        }
        return len1 - len2;
    }
}