using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Xml;
using System.IO.Packaging;
using ClosedXML.Excel;
using System.Text;

// Must Reference DocumentFormat.OpenXml, Microsoft.Office.Interop.Excel and WindowsBase to use this file

/*********** Basic Microsoft overview **************
The basic structure of a SpreadsheetXML document consists of the Sheets and Sheet elements, which reference the worksheets in the workbook. 
For example if the  workbook has two sheets (tabs) the XML is similar to:

<?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
<workbook xmlns=http://schemas.openxmlformats.org/spreadsheetml/2006/main xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
        <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
    </sheets>
</workbook>

A separate XML file is created for each Worksheet.  The worksheet XML files contain SheetData. sheetData represents the cell table and contains one or more Row elements. 
A row contains one or more Cell elements. Each cell contains a CellValue element that represents the value of the cell. For example:

<?xml version="1.0" encoding="UTF-8" ?> 
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1">
                <v>100</v> 
            </c>
        </row>
    </sheetData>
</worksheet>

*** XML to Class mapping ***
  sheetData = SheetData
  row       = Row
  c         = Cell
  v         = CellValue

*/

/// <summary> Wrapper class for reading/writting Excel documents in openXML format (.XLSX) </summary>
/// <example><code>
/// var excelObj = new Excel(@"C:\DEV\NETLIST.xlsx");
/// string cellVal = excelObj.ReadCellValue("Engineering Notes", "D8");
/// </code></example>
public class Excel : object  // No inheritance for now
{
    // Properties
    private readonly string excelFilePath;
    public SpreadsheetDocument excelDoc;
    public WorkbookPart workbookPart = null;
    private Workbook workbook = null;
 
    // Constructors
    //private Excel(); // No public constructior with NO parameters
    public Excel(string filePath, bool writeable=true)
    {
        // Do we want to use m_ for private member properties??
        excelFilePath = filePath;
        try
        {
            excelDoc = SpreadsheetDocument.Open(excelFilePath, writeable);
            workbookPart = excelDoc.WorkbookPart;
            workbook = workbookPart.Workbook;
        }
        catch
        {
            System.Threading.Thread.Sleep(500);
            OpenSettings openSettings = new OpenSettings();
            MarkupCompatibilityProcessSettings mySettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessLoadedPartsOnly, DocumentFormat.OpenXml.FileFormatVersions.Office2010);

            var compat = openSettings.MarkupCompatibilityProcessSettings;
            openSettings.MarkupCompatibilityProcessSettings = mySettings;
            excelDoc = SpreadsheetDocument.Open(excelFilePath, writeable, openSettings);
            workbookPart = excelDoc.WorkbookPart;
            workbook = workbookPart.Workbook;
        }
    }
    ~Excel() // Destructor
    {
        if (excelDoc != null)
        {
            //excelDoc.Close(); // This is sometimes not open. Dispose will close if it is
            excelDoc.Dispose();
        }
    }

    public void Close()
    {
        excelDoc.Close();
    }

// *** PRIVATE Methods

    private Cell GetCell(string sheetName, string cellCoordinates)
    {
        WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
        return GetCell(worksheetPart.Worksheet, cellCoordinates);
    }
    // Get a Cell object given a Worksheet object and cell coordinates (B6)
    private static Cell GetCell(Worksheet worksheet, string colAndRow)
    {
        return worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == colAndRow);
    }
    // Given a worksheet and a row index, return the row object
    private static Row GetRow(Worksheet worksheet, uint rowIndex)
    {
        return worksheet.GetFirstChild<SheetData>().Elements<Row>().First(r => r.RowIndex == rowIndex);
    }

    private string GetCellValue(SpreadsheetDocument document, Cell cell)
    {
        SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
        if (cell.CellValue == null) return "";
        string value = cell.CellValue.InnerXml;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;

        return value;
    }

    private string ReadCell(Cell cell)
    {
        if (cell == null || cell.ChildElements.Count == 0) return null; // Cell has no settable parts so get stop here

        var cellValue = cell.CellValue;
        var value = (cellValue == null) ? cell.InnerText : cellValue.Text;
        string cellAddress = cell.CellReference;

        /*
        var links = workbookPart.WorksheetParts.First().Worksheet.Elements<Hyperlinks>().First(); // Collection is all under [0] for some reason
        foreach (Hyperlink theLink in links)
        {
            if (theLink.OuterXml.Contains(cellAddress))
            {
                var test = theLink.InnerXml;
                string test2 = theLink.Parent.ToString();
            }

            var theCell = theLink.Descendants<Cell>().Where(c => c.CellReference == cellAddress).FirstOrDefault();
            if (theCell == null) continue;
            var innerXML = theCell.InnerXml;
            return ReadCell(theCell);
        }

        int cnt = workbookPart.SharedStringTablePart.SharedStringTable.Elements<MergeCells>().Count();

        if (cnt > 0)
        {
            var merged = workbookPart.WorksheetParts.First().Worksheet.Elements<MergeCells>().First(); // Collection is all under [0] for some reason
            foreach (MergeCell mc in merged)
            {
                if (mc.Reference.InnerText.StartsWith("A23"))
                {
                    var test = mc.InnerXml;
                    string test2 = mc.Reference;
                }

                var theCell = mc.Descendants<Cell>().Where(c => c.CellReference == cellAddress).FirstOrDefault();
                if (theCell == null) continue;
                var innerXML = theCell.InnerXml;
                return ReadCell(theCell);
            }
        }
        */
        if (cell.DataType == CellValues.SharedString)
        {
            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
            var cellObject = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)];
            var returnValue = (cellObject.InnerText != null) ? cellObject.InnerText : cellObject.InnerXml;

            return returnValue;
            /* TESTING FOR ADVANCED PROPERTIES LIKE FORMULAS
            string cellAddress = cell.CellReference;
            var formula = cell.CellFormula;
            if (formula != null)
            {
                var t = formula.Text;
                var t2 = formula.InnerText;
                var outer = formula.OuterXml;
            }
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(workbookPart.Workbook.OuterXml);

            // Create an XmlNamespaceManager to resolve namespaces.
            NameTable nt = new NameTable();
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(new NameTable());
            nsmgr.AddNamespace("dft", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"); // yes Microsoft still uses 2006 for the main namespace
            nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            string searchString = "/dft:worksheet/dft:hyperlinks/dft:hyperlink[@ref='D5']/@r:id";

            var node = doc.SelectSingleNode(searchString, nsmgr);
            if (node != null)
            {
                var attr = node.Attributes.GetNamedItem("id");
                var relID = attr.Value;
                var hyperlink = stringTablePart.GetExternalRelationship(relID);
            }

             * XmlNode linkNode = _cellElement.OwnerDocument.SelectSingleNode(searchString, workbookPart.Na_xlWorksheet.NameSpaceManager);
            if (linkNode != null)
            {
                XmlAttribute attr = (XmlAttribute)linkNode.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships);
                if (attr != null)
                {
                    string relID = attr.Value;
                    // now use the relID to lookup the hyperlink in the relationship table
                    PackageRelationship relationship = _xlWorksheet.Part.GetRelationship(relID);
                    _hyperlink = relationship.TargetUri;
                }
            }
            var testType = stringTablePart.Parts.GetType().ToString();
            // If it is a URL we can use this:
            var checkType = stringTablePart.HyperlinkRelationships.ElementAtOrDefault(6);
            HyperlinkRelationship hr = workbookPart.HyperlinkRelationships.Where(a => a.Id == "rId6").FirstOrDefault();
            if (hr != null)
            {
                var test = hr.Uri.AbsoluteUri;
                var test2 = hr.Uri.ToString();
            }
            SharedStringItem item = GetSharedStringItemById(workbookPart, Int32.Parse(value));
            if (item.Text != null)
            {
                return item.Text.Text;
            }
            else if (item.InnerText != null)
            {
                return item.InnerText;
            }
            else if (item.InnerXml != null)
            {
                return item.InnerXml;
            }
            // string link = workbookPart.HyperlinkRelationships.ElementAtOrDefault(Int32.Parse(value)).Uri.ToString();
            value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            */
        }
        return value.Trim();
    }

    public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
    {
        return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
    }

    private static string GetColumnName(string cellReference)
    {
        var regex = new Regex("[A-Za-z]+"); // Just the letter portion of AB99
        var match = regex.Match(cellReference);
        return match.Value;
    }
    public static string GetRowIndex(string cellReference)
    {
        return new Regex(@"\d+").Match(cellReference).Value;
    }


// *** PUBLIC Methods

    /// <summary> This returns the value of the cell even if it is a linked shared cell </summary>
    /// <param name="sheetName"></param>
    /// <param name="cellCoordinates"></param>
    /// <returns>value of string</returns>
    /// <example><code>
    /// var excelObj = new Excel(@"C:\DEV\NETLIST.xlsx");
    /// string cellVal = excelObj.ReadCell("Engineering Notes", "D8");
    /// </code></example>
    public string GetCellValue(string sheetName, string cellCoordinates)
    {
       Cell cell = GetCell(sheetName, cellCoordinates);
       return ReadCell(cell);
    }

    public string GetCell(WorksheetPart worksheetPart, string cellCoordinates)
    {
        Cell cell = GetCell(worksheetPart.Worksheet, cellCoordinates);
        return ReadCell(cell);
    }

    /// <summary>Given a DataTable and a column to search returns the row number or -1 if not found</summary>
    public static int FindRowNumber(DataTable oTable, string columnNameOrNumber, string stringToFind, bool exactMatch = true)
    {
        int i = 1;
        foreach (DataRow row in oTable.Rows)
        {
            i++;
            if (exactMatch)
                if (row[columnNameOrNumber].ToString() == stringToFind) return i;
                else
                    if (row[columnNameOrNumber].ToString().Contains(stringToFind)) return i;
        }
        //else
        return -1; // to tell them NOT found
    }
    /// <summary> Given a sheetname gets the WorksheetPart. Defaults to First sheet</summary>
    /// <param name="sheetName"></param>
    /// <returns>WorksheetPart</returns>
    public WorksheetPart GetWorksheetPart(string sheetName="")
    {
        try
        {
            string ID = workbook.Descendants<Sheet>().First(s => s.Name == sheetName).Id;
            return workbookPart.GetPartById(ID) as WorksheetPart;
        }
        catch (Exception)
        {
            // Get the first sheet. Tried to get the first visible but s.State is always null
            Sheet firstSheet = workbookPart.Workbook.Descendants<Sheet>().First(); //s => s.State == SheetStateValues.Visible);
            Worksheet firstWorksheet = ((WorksheetPart)workbookPart.GetPartById(firstSheet.Id)).Worksheet;

            string name = workbookPart.Workbook.Descendants<Sheet>().ElementAt(0).Name;
            return firstWorksheet.WorksheetPart;
        }
    }

    /// <summary> Given a sheetname gets the WorksheetPart. Defaults to First sheet</summary>
    /// <param name="sheetName"></param>
    /// <returns>WorksheetPart</returns>
    public WorksheetPart GetWorksheetPart(string[] sheetName)
    {
        try
        {
            string ID = workbook.Descendants<Sheet>().First(s => ((IList<string>)sheetName).Contains(s.Name)).Id;
            return workbookPart.GetPartById(ID) as WorksheetPart;
        }
        catch (Exception)
        {
            // Get the first sheet. Tried to get the first visible but s.State is always null
            Sheet firstSheet = workbookPart.Workbook.Descendants<Sheet>().First(); //s => s.State == SheetStateValues.Visible);
            Worksheet firstWorksheet = ((WorksheetPart)workbookPart.GetPartById(firstSheet.Id)).Worksheet;

            string name = workbookPart.Workbook.Descendants<Sheet>().ElementAt(0).Name;
            return firstWorksheet.WorksheetPart;
        }
    }

    /// <summary>Update a single cell by Column/Row coordinates  </summary>
    /// <param name="filename"></param>
    /// <param name="sheetName"></param>
    /// <param name="cellCoordinates"></param>
    /// <param name="newValue"></param>
    public void UpdateCellValue(string sheetName, string cellCoordinates, object newValue)
    {
        // tell Excel to recalculate formulas next time it opens the doc
        workbook.CalculationProperties.ForceFullCalculation = true;
        workbook.CalculationProperties.FullCalculationOnLoad = true;

        WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
        Cell cell = GetCell(worksheetPart.Worksheet, cellCoordinates);
        cell.CellValue = new CellValue(newValue.ToString());
        cell.DataType = CellValues.String; // In-case it was some Shared string (with just a reference # in the cell)
        worksheetPart.Worksheet.Save();
    }

    /// <summary> Creates a new sheet/tab with the name passed and returns the WorksheetPart object to it</summary>
    /// <param name="sheetName"></param>
    /// <returns>WorksheetPart</returns>
    public WorksheetPart CreateNewSheet(string sheetName)
    {
        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0)
        {
            sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }

        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);

        return GetWorksheetPart(sheetName);
    }

    /// <summary> Returns all sheets/tab as a Dictionary collection</summary>
    /// <param name="worksheetPart"></param>
    /// <returns></returns>
    public Dictionary<string, string> GetSheetAttributes(WorksheetPart worksheetPart)
    {
        Dictionary<string, string> keyArray = new Dictionary<string,string>(); // Empty collection

        var sheet = workbookPart.Workbook.Sheets;
        int count = sheet.Elements<Sheet>().Count();

        Sheets sheets = workbook.GetFirstChild<Sheets>();
        int count2 = sheets.Elements<Sheet>().Count();

        foreach (var attr in worksheetPart.Worksheet.SheetProperties.GetAttributes())
        {
            keyArray.Add(attr.LocalName, attr.Value);
        }
        return keyArray;
    }


    public DataTable GetAllSheets()
    {
        DataTable myTable = new DataTable();
        // This gets all sheets (GetFirstChild is just for the workbook "part")
        Sheets sheets = workbook.GetFirstChild<Sheets>();
        int count = sheets.Elements<Sheet>().Count();
        if (count == 0) return myTable; // No sheets so return an empty table

        // First load attribute names as the column names in the DataTable
        var firstSheet = sheets.FirstOrDefault();
        foreach (var attr in firstSheet.GetAttributes())
        {
            myTable.Columns.Add(attr.LocalName);
        }

        foreach (var sheet in workbook.Sheets)
        {
            DataRow row = myTable.NewRow(); // Must be inside the loop or you will get a message saying the row already belongs to the table
            foreach (var attr in sheet.GetAttributes()) { row[attr.LocalName] = attr.Value; }
            myTable.Rows.Add(row);
        }
        return myTable;
    }

    /// <summary>Forces a refresh of all calculated columns</summary>
//    public void Refresh()
//    {
//        var excelApp = new Microsoft.Office.Interop.Excel.Application();
//        var workbook = excelApp.Workbooks.Open(excelFilePath);
//        workbook.Close(true);
//        excelApp.Quit();
//    }

    /// <summary>This uses Interop. Use the other GetCell methods whenever possible </summary>
    /// <param name="inFileName"></param>
    /// <param name="col"></param>
    /// <param name="row"></param>
    /// <param name="optionalSheet"></param>
    /// <returns>value of cell as string</returns>
/*    public static string GetCellValue(string inFileName, string col, string row, string optionalSheet="")
    {
        var excelApp = new Microsoft.Office.Interop.Excel.Application();
        excelApp.DisplayAlerts = false;
        excelApp.Visible = false;
        Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(inFileName);
        int count = workbook.Worksheets.Count;
        var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
        if (optionalSheet != "")
        {
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[optionalSheet];        }
        string sheetName = worksheet.Name;
        Microsoft.Office.Interop.Excel.Range excelRange = worksheet.get_Range(col + row);
        string retVal = excelRange.Text.ToString();
        if (excelRange.Hyperlinks.Count > 0) // Any links
            retVal = "<a href='" + excelRange.Hyperlinks[1].Address + "'>" + retVal + "</a>";

        workbook.Close(0);
        excelApp.Quit();

        Marshal.FinalReleaseComObject(workbook);
        Marshal.ReleaseComObject(excelApp.Workbooks);
        Marshal.FinalReleaseComObject(excelApp);

        //GC.Collect();
        //GC.WaitForPendingFinalizers();
        return retVal;
    }
*/
    /// <summary>Reads a DataTable like an Excel spreadsheet with A3 "coordinates meaning column 0 and row 2 (zero-based). Only works for A-Z columns</summary>
    /// <param name="theTable"></param>
    /// <param name="rowNum"></param>
    /// <param name="columnLetter"></param>
    /// <returns></returns>
    public static string ReadTable(DataTable theTable, string colAndRow)
    {
        string columnLetter = colAndRow.Substring(0, 1);
        int rowNum = Convert.ToInt32(colAndRow.Substring(1));
        int colNumber = ColLetterToNumber(columnLetter);
        return theTable.Rows[rowNum-1][colNumber-1].ToString();
    }

    /// <summary>Uses Interop to convert a .XLS to an .XLSX file</summary>
    /// <param name="inFileName"></param>
    /// <param name="outFileName"></param
/*
    public static void ConvertXlsToXLSX(string inFileName, string outFileName)
    {
        var excelApp = new Microsoft.Office.Interop.Excel.Application();
        excelApp.DisplayAlerts = false;
        excelApp.Visible = false;
        var workbook = excelApp.Workbooks.Open(inFileName);
        workbook.SaveAs(outFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);

        workbook.Close();
        excelApp.Quit();

        Marshal.FinalReleaseComObject(workbook);
        Marshal.FinalReleaseComObject(excelApp);
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    /// <summary>Uses Interop to replace all occurances just like a Find and Replace of all cells for a value</summary>
    /// <param name="filename"></param>
    /// <returns>boolean indicating if updates were actually made</returns>
    public static bool ReplaceInExcel(string filename)
    {
        var excelApp = new Microsoft.Office.Interop.Excel.Application();
        //xlapp.Application.Interactive = true;  // This will let you know when a find and replace operation does not find the text you think it should
        //xlapp.Application.UserControl = true;  // This is no good on a server or a dialog will hide on the console of say BOWASMPD01
        var workbooks = excelApp.Workbooks;
        var workbook = workbooks.Open(filename);
        var sheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Worksheets[1];

        string sheetName = sheet.Name;
        //Range excelCell = (Range)sheet.get_Range("F3", "F3");

        Microsoft.Office.Interop.Excel.Range myRange = sheet.Cells.Find("moss.micron.com");
        bool didItUpdate = false;
        if (myRange != null)
        {
            didItUpdate = true;
            string text = myRange.Text.ToString();

            // *Beware* XlLookAt.xlWhole means the whole cell must match so use xlPart
            sheet.Cells.Replace(What: "moss.micron.com", Replacement: "collab.micron.com", LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                    SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, MatchCase: true, SearchFormat: false, ReplaceFormat: false);

            //sheet.Hyperlinks.Add(myRange, "http://google.com/", Type.Missing, "Google", "Google");

            workbook.Save(); // Write back to disk
        }
        workbook.Close();
        excelApp.Quit();
        Marshal.FinalReleaseComObject(sheet);
        Marshal.FinalReleaseComObject(workbook);
        Marshal.FinalReleaseComObject(workbooks);
        Marshal.FinalReleaseComObject(excelApp);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        return didItUpdate;
    }

    /// <summary>Returns a DataTable with the data from the Sheet name passed. First row is used for column names. SheetToDataTable does the same thing but uses the Microsoft OLEDB driver </summary>
    /// <param name="sheetName"></param>
    /// <returns>DataTable</returns>
    public DataTable ExcelToDataTable(string sheetName)
    {
        DataTable dt = new DataTable();
        WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
        SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;

        Row lastRow = worksheetPart.Worksheet.Descendants<Row>().LastOrDefault();
        Row firstRow = worksheetPart.Worksheet.Descendants<Row>().FirstOrDefault();
        if (firstRow != null)
        {
            foreach (Cell c in firstRow.ChildElements)
            {
                string value = ReadCell(c);
                dt.Columns.Add(value);
            }
        }
        if (lastRow != null)
        {
            for (int i = 2; i <= lastRow.RowIndex; i++)
            {
                DataRow dr = dt.NewRow();
                bool empty = true;
                Row row = worksheetPart.Worksheet.Descendants<Row>().Where(r => i == r.RowIndex).FirstOrDefault();
                int j = 0;
                if (row != null)
                {
                    foreach (Cell c in row.ChildElements)
                    {
                        string value = ReadCell(c);
                        if (!string.IsNullOrEmpty(value) && value != "") empty = false;
                        dr[j] = value;
                        j++;
                        if (j == dt.Columns.Count) break;
                    }

                    if (empty) break;
                    dt.Rows.Add(dr);
                }
            }
        }

        return dt;
    }
*/
    /// <summary>Returns a DataTable with the data from the Sheet name passed. Assumes 1st row is the header/column names </summary>
    /// <param name="sheetName"></param>
    /// <returns>DataTable</returns>
    public static DataTable SheetToDataTable(string excelFilePath, string sheetName="", bool firstRowHeaders=true)
    {
        string headers = "NO";
        if (firstRowHeaders) headers = "YES";

    	//string dsn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 8.0;HDR=" + headers + "\"";
        string dsn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 8.0;HDR=" + headers + "\"";

        if (excelFilePath.ToLower().Contains(".xlsx"))
        {
            dsn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=" + headers + "\"";
            //dsn = "Provider=Microsoft.ACE.OLEDB.15.0;Data Source=" + excelFilePath + ";Extended Properties=Excel 15.0 Xml;IMEX=1;HDR=" + headers + "\"";
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

    /// <summary>Returns a DataTable with the data from the Sheet name passed. Assumes 1st row is the header/column names </summary>
    /// <param name="sheetName"></param>
    /// <returns>DataTable</returns>
    public static string GetCellValue(DataTable oTable, string columnLetter, int row, bool firstRowHeaders = true) // Only A through Z for now
    {
        int colNumber = ColLetterToNumber(columnLetter);
        row--; // DataTable is zero based but looking at Excel is 1 based so subtract 1
        if (firstRowHeaders) row--;
        string results = oTable.Rows[row][colNumber].ToString();
        return results;
    }

    /// <summary>Convert A to 1 and B to 2, etc.</summary>
    public static int ColLetterToNumber(string columnLetter)
    {
        // Rian: TO handle double letter columns AA-AZ you can use:
        return columnLetter.ToUpper().Aggregate(0, (column, letter) => 26 * column + letter - 'A' + 1);

        char myChar = Convert.ToChar(columnLetter.ToUpper()); // Need upper case column letters A for 

        // ASCII A = 65
        return System.Convert.ToInt32(myChar) - 65;  // ASCII 65=A so if we want A to be column 1 subtract 64
    }

    public static string NumberToColLetter(int columnNumber)
    {
        int dividend = columnNumber;
        string columnName = String.Empty;
        int modulo;

        while (dividend > 0)
        {
            modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo).ToString() + columnName;
            dividend = (int)((dividend - modulo) / 26);
        } 

        return columnName;
    }

    public static void UpdateCell(string fileName, string sheetName, string column, int row, string valueToSet)
    {
        string dsn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=NO\"";
        if (fileName.ToLower().Contains(".xlsx"))
        {
            dsn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'"; // IMEX=1
        }
        OleDbConnection dbConnection = new OleDbConnection(dsn);
        dbConnection.Open();

        if (sheetName == "") sheetName = GetFirstSheetName(dbConnection);
        if (!sheetName.EndsWith("$")) sheetName += "$";

        // EX: UPDATE [Sheet1$A2:A2] SET F1='TestValue1'
        string colRow = column + row.ToString();
        string sql = "Update [" + sheetName + colRow + ":" + colRow + "] set F1 = '" + valueToSet.Replace("'", "''") + "'";

        OleDbCommand cmd = new OleDbCommand(sql, dbConnection);
        //cmd.Parameters.AddWithValue("@p1", valueToSet);
        cmd.ExecuteNonQuery();
        dbConnection.Close();
    }

    public static string GetFirstSheetName(OleDbConnection dbConnection)
    {
        DataTable dbSchema = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        if (dbSchema == null || dbSchema.Rows.Count < 1) throw new Exception("No Sheet found in .XLS file.");
        string sheetName = dbSchema.Rows[0]["TABLE_NAME"].ToString();
        //dbConnection.Close();
        return sheetName.Replace("'", "");
    }


    public static DataTable RangeToDataTable(IXLRange range, bool firstLineHeadings = true, bool convertRichTextToHTML = true)
    {
        if (range.IsMerged())
        {
            range = range.Unmerge(); // Unmerge any merged cells
        }

        DataTable dt = new DataTable();
        // First just do headers (1st row in Range)
        int columnCnt = range.ColumnCount();
        for (int k = 1; k <= columnCnt; k++)
        {
            var name = range.Row(1).Cell(k).Address.ColumnLetter;
            if (firstLineHeadings) name = range.Row(1).Cell(k).GetString();
            try {
				dt.Columns.Add(name);
			} catch (Exception e) {
				dt.Columns.Add(name + "A");
			}
        }
        dt.Columns.Add("FromExcelRow");

        bool firstRow = true;
        int rowNum = 0;
        foreach (var myrow in range.RowsUsed()) // RowsUsed seems to also make merged called all report the same value (which we want)
        {
            if (firstRow && firstLineHeadings) { firstRow = false; continue; }
            if (myrow.IsEmpty()) continue;
            DataRow workRow = dt.NewRow();
            for (int col = 0; col < columnCnt; col++)
            {
                IXLCell cellObj = myrow.Cell(col + 1); // not 0 based so we need to add 1
                var cellValue = "";
                if (convertRichTextToHTML)
                {
                    if (cellObj.Address.ColumnLetter == "Q") { int stop = 1; }
                    cellValue = RichTextToHTML(cellObj); // Even if there is no RichText this will get the value in the cell
                }
                else
                    cellValue = cellObj.GetString();
/*
                if (cellObj.Address.ColumnLetter == "H" && cellValue != "")
                {
                    cellValue = RichTextToHTML(cellObj);
                    int findRed = cellValue.IndexOf("color: Red") + 10;
                    if (findRed > 0) cellValue = cellValue.Substring(cellValue.IndexOf(':', findRed) + 1).Trim();
                    cellValue = Regex.Replace(cellValue, "<.*?>", "");
                }
*/

                if (cellObj.IsMerged() && cellValue.ToString() == "")  // then get last cell (yes we assume merging up and down here)
                {
                    int rowNo = cellObj.Address.RowNumber; // This is the row # within the sheet
                    if (rowNum > 0) workRow[col] = dt.Rows[rowNum - 1][col].ToString(); // Look "up" one 1 in same column
                }
                else
                    workRow[col] = cellValue; // 0 based in DataRow object for [] indexing
            }
            workRow[columnCnt] = myrow.RowNumber(); // 0 based in DataRow object for [] indexing
            dt.Rows.Add(workRow);
            rowNum++;
        }
        return dt;
    }

    public static DateTime FromExcelSerialDate(object serialDate)
    {
        int excelNumber = int.Parse(serialDate.ToString());
        if (excelNumber > 59) excelNumber -= 1; //Excel/Lotus 2/29/1900 bug   
        return new DateTime(1899, 12, 31).AddDays(excelNumber);
    }

    public static string RichTextToHTML(IXLCell cell)
    {
        var cellValue = cell.GetString();
        var checkForDate = cell.GetFormattedString();
        if (checkForDate.Contains("-yy") && checkForDate.Contains("[$-"))
            cellValue = FromExcelSerialDate(cellValue).ToShortDateString();
        else
            if (cell.HasHyperlink && cell.Hyperlink.IsExternal) cellValue = "<a href='" + cell.Hyperlink.ExternalAddress + "'>" + cell.Value + "</a>";
        if (!cell.HasRichText) return cellValue;

        StringBuilder myString = new StringBuilder();
        foreach (var richText in cell.RichText)
        {
            myString.Append("<span style='"); // Start SPAN and CSS here then add to it below
            // Now look for attributes that make up the Richtext
            if (richText.Bold) myString.Append("font-weight: bold;");
            if (richText.Strikethrough) myString.Append("text-decoration:line-through;");
            if (richText.Italic) myString.Append("font-style: italic;");
            if (richText.Underline == XLFontUnderlineValues.Single) myString.Append("text-decoration: underline;");
            if (richText.FontName != "Arial") myString.Append("font-face: " + richText.FontName + ";");
            if (richText.FontColor.Color.IsKnownColor) myString.Append("color: " + richText.FontColor.Color.Name + ";");
            if (richText.FontSize != 10) myString.Append("font-size: " + richText.FontSize + ";");
            myString.Append("'>"); // End the SPAN
            if (richText.Text.Contains("http")) myString.Append("<a href='" + richText.Text + "'>");

            myString.Append(richText.Text.Replace("\r\n", "<br>"));
            if (richText.Text.Contains("http")) myString.Append("</a>");
            myString.Append("</span>");
        }
        return myString.ToString();
    }

    public static DataTable WorkSheetToDataTable(string filename, bool firstLineHeadings = true)
    {
        var wb = new XLWorkbook(filename);
        var ws = wb.Worksheet(1); // Assume first sheet
        return WorkSheetToDataTable(ws, firstLineHeadings);
    }

    public static DataTable WorkSheetToDataTable(IXLWorksheet worksheet, bool firstLineHeadings = true)
    {
        var dt = new DataTable();
        var columnCnt = worksheet.ColumnsUsed().Count();
        var RowsQty = worksheet.RowsUsed().Count();

        for (int k = 1; k <= columnCnt; k++)
        {
            var name = worksheet.Row(1).Cell(k).Address.ColumnLetter;
            if (firstLineHeadings) name = worksheet.Row(1).Cell(k).GetString();
            dt.Columns.Add(name);
        }
        dt.Columns.Add("FromExcelRow");
        bool firstRow = true;
        foreach (var myrow in worksheet.RowsUsed())
        {
            if (firstRow && firstLineHeadings) { firstRow = false; continue; }
            DataRow workRow = dt.NewRow();
            for (int i = 1; i <= columnCnt; i++)
            {
                var cellObj = myrow.Cell(i);
                var cellAddr = cellObj.Address.ToString();
                var cellValue = cellObj.Value;

                var checkForDate = cellObj.GetFormattedString();
                if (checkForDate.Contains("-yy") && checkForDate.Contains("[$-"))
                    cellValue = FromExcelSerialDate(cellValue);
                else
                {
                    if (cellObj.HasHyperlink && cellObj.Hyperlink.IsExternal) cellValue = "<a href='" + cellObj.Hyperlink.ExternalAddress + "'>" + cellObj.Value + "</a>";
                    if (cellObj.HasRichText) cellValue = RichTextToHTML(cellObj); // If the link is formatted this overwrites with the prettier link
                }
                workRow[i - 1] = cellValue;
            }
            workRow[columnCnt] = myrow.RowNumber(); // 0 based in DataRow object for [] indexing

            dt.Rows.Add(workRow);
        }
        return dt;
    }

    // Return blank if not found
    public static string FindFirstMatch(IXLWorksheet ws, string findString, int startRow = 1, bool exactMatch = false)
    {
        var dt = Excel.WorkSheetToDataTable(ws, firstLineHeadings: false);
        var lastColumn = dt.Columns.Count;
        foreach (DataRow theRow in dt.Rows)
        {
            if (Convert.ToInt32(theRow["FromExcelRow"]) < Convert.ToInt32(startRow)) continue;
            var colCount = 1;
            foreach (DataColumn theCol in dt.Columns)
            {
                var theValue = theRow[theCol].ToString();
                if (theValue.Contains(findString)) return Excel.NumberToColLetter(colCount) + theRow[lastColumn - 1];
                colCount++;
            }
        }
        return "";
    }

    // Return blank if not found. Calls above method but must open file each time. Good for a couple of reads. For more use above so the file is not open and closed as much
    public static string FindInSheet(string filename, string sheetName, string findString, bool exactMatch=false)
    {
        var wb = new XLWorkbook(filename);
        var ws = wb.Worksheet(sheetName);
        return FindFirstMatch(ws, findString, 1, exactMatch);
    }
}