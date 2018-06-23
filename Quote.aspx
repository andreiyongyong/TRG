<%@ PAGE language=C# debug="true" EnableSessionState="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="System.Xml.XPath" %>
<!--#include file="headerNEW.aspx"-->
<link rel="stylesheet" href="css/style.css">
<link rel="stylesheet" href="css/productSuggestions.css">


<link rel="stylesheet" type="text/css" media="screen" href="css/BootstrapRian.css" />
<style>
  body {text-align: left; line-height: 1.0;}
  label {font-family: Calibri;font-size: 110%;display: inline-block;width: 150px;}
  form {padding: 0;}
</style>
<script>
$(function() {
  $('#Email').prop('type', 'email'); // Can't set in markup when runat="server" is present
});
</script>
<br>
<asp:Literal id="SQL" runat="server" />
<img class="pull-right" src='/images/products/detail/<%= Util.Request("ProductCode") %>.jpg'>
<form method="GET" class="panel panel-info" onsubmit="  var minQty = <asp:Literal id="MinimumQty" runat="server" />;
  var qty1 = $('#Quantity1').val();
  if (qty1 < minQty) {alert('Minimun order is ' + minQty);return false;}
  return true;" style="width: 600px; margin-left: 250px; border: 1px solid skyblue;">
<input type="hidden" name="ProductCode" value="<%= Util.Request("ProductCode") %>">
  <div class="panel-heading">
  <button class="btn btn-primary center pull-right" style="margin-top: -5px;" onclick="window.close();">Close</button>

  Price Calculator for Product Code: <asp:Literal id="ProductCode" runat="server" /></div>
  <div class="panel-body">
    <h3 style="color: blue; margin: -10px 0 20px 0; text-align: center;"><asp:Literal id="ProductName" runat="server" /></h3>
    <label>Ship to Address:</label> City: <input type="text" id="shipcity" class="input-sm" style="width: 100px;" runat="server" />
      State: <input type="text" id="shipstate" class="input-sm" style="width: 50px;" runat="server" />
      Postcode: <input type="text" id="postcode" class="input-sm" style="width: 50px;" runat="server" /><br>
    <label>Email:</label> <input id="Email" size="50" class="input-sm" placeholder="(optional)" runat="server" /><br><br>

    <label>Quantity 1:</label> <input type="text" id="Quantity1" required="true" class="input-sm" runat="server" /> <small class="text-muted"> <asp:Literal id="MinQty" runat="server" /></small><br>
    <label>Quantity 2:</label> <input type="text" id="Quantity2" placeholder="(optional)" class="input-sm" runat="server" /><br>
    <label>Quantity 3:</label> <input type="text" id="Quantity3" placeholder="(optional)" class="input-sm" runat="server" /><br>
    
    <asp:Literal id="Attachment" runat="server" />
    <label>Time Frame:</label> <select name="TimeFrame" class="input-sm" required="true"><option value=""></option>
      <asp:Literal id="TimeFrame" runat="server" />
    </select><br>
    <label>Decorations:</label> <select name="Decorations" class="input-sm">
      <asp:Literal id="Decorations" runat="server" />
    </select><br>
    <label>Number of Colours:</label> <select id="NumberColours" class="input-sm" runat="server">
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4+ (Please contact us) </option>
    </select><br>
    <label>Setup Costs:</label> 
    <input type="radio" name="setupCost" checked="1" value="<asp:Literal id="SetupCost" runat="server" />">$<asp:Literal id="SetupCost2" runat="server" /> Initial 
    <input type="radio" name="setupCost" value="<asp:Literal id="RepeatCost" runat="server" />" style="margin-left:10px">$<asp:Literal id="RepeatCost2" runat="server" /> Repeat<br><br>

    <asp:Literal id="QuoteString" runat="server" />
    <asp:Literal id="PrintSize" runat="server" /><br>
    Qty per box: <asp:Literal id="QtyPerBox" runat="server" /> Weight per box: <asp:Literal id="WeightPerBox" runat="server" /> kg<br>
  </div>

  <button class="btn btn-primary center" style="margin-top: 10px;" type="submit">Get Quote with Shipping</button>
</form>
<div class="panel panel-info center" style="width: 600px">
  <asp:Literal id="PriceColumn" runat="server"/>
</div>
<xmp style="margin-top:150px;font-size: 80%; color: maroon;">
  <asp:Literal id="Debug" runat="server"/>
</xmp>
<script runat="server">
void Page_Load()
{
  Util.DefaultInputs();

  var productCode = Util.Request("productCode");
  ProductCode.Text = productCode;
	var sql = @"
    SELECT PrintCodes.SetupCost as SetupCost2, PrintCodes.RepeatSetup as RepeatSetup2, Product.* FROM Product
    RIGHT OUTER JOIN PrintCodes ON PrintCodes.PrintCodeID = Product.PrintCodeID
    WHERE productCode='" + productCode + "'";
	var db = new DBHelper(sql);
	if (db.Count == 0) Util.AlertAndEnd(productCode + " not found.");

  //SQL.Text = db.MakeHTMLTable();
  var attachments = new DBHelper();

  if (productCode.StartsWith("LAN"))
  {
    attachments = new DBHelper(@"
      SELECT '<img src=' + Thumbnail + '>' as Thumbnail, AttachmentName as Attachment, Price100 as [Price 100+], Price250 as [Price 250+], Price500 as [Price 500+], Price1000 as [Price 1000+] 
      FROM LanyardAttachments
      WHERE AttachmentName NOT IN ('')
      ORDER BY AttachmentName");
    
    Attachment.Text = @"<label>Attachment 1:</label> <select name='LanyardAttachment'>";
     
    PriceColumn.Text = attachments.MakeHTMLTable();
    foreach (DataRow row in attachments.Rows) {
      var selected = "";
      if (Util.Request("LanyardAttachment") == row["Attachment"].ToString() ) selected = " SELECTED";
      Attachment.Text += "<option" + selected + ">" + row["Attachment"] + "</option>\n";
    }
    Attachment.Text += "</select><br>";
    Attachment.Text += @"<label>Attachment 2:</label> <select name='LanyardAttachment2'><option></option>";
    foreach (DataRow row in attachments.Rows) {
      var selected = "";
      if (Util.Request("LanyardAttachment2") == row["Attachment"].ToString() ) selected = " SELECTED";
      Attachment.Text += "<option" + selected + ">" + row["Attachment"] + "</option>\n";
    }
    Attachment.Text += "</select><br>";    
  }
	var id = db.Get("ProductId");
  var prefixProductCode = Util.Left(productCode, 3);
  ProductName.Text = db.Get("ProductName");
  SetupCost.Text = SetupCost2.Text = db.Get("SetupCost");
  RepeatCost.Text = RepeatCost2.Text = db.Get("RepeatSetup");
  if (SetupCost.Text == "") SetupCost.Text = SetupCost2.Text = db.Get("SetupCost2");
  if (RepeatCost.Text == "") RepeatCost.Text = RepeatCost2.Text = db.Get("RepeatSetup2");
  var qty = Util.Request("Quantity1"); if (qty == "") qty = "100";
  var quantity = Util.ToInt(qty);
  var quantity2 = Util.ToInt(Util.Request("Quantity2"));
  var quantity3 = Util.ToInt(Util.Request("Quantity3"));
  sql = @"
      SELECT * FROM PrintQtyHeadings WHERE ProductID=" + id + " ORDER BY RowId";
	db = new DBHelper(sql);
  var minQty = "100";
  if (db.Count > 0) minQty = db.Get("Qty");
  int i = 0, priceRange = 0, priceRange2 = 0, priceRange3 = 0;
  var allQty = "";
  foreach (DataRow row in db.Rows) {
    var rowQty = row["Qty"].ToString().Replace("+", "");
    allQty += rowQty + ", ";
    if (quantity >= Util.ToInt(rowQty))  priceRange = i;
    if (quantity2 >= Util.ToInt(rowQty)) priceRange2 = i;
    if (quantity3 >= Util.ToInt(rowQty)) priceRange3 = i;
    i++;
  }
  //if (priceRange == 0) Util.AlertAndEnd("priceRange is not found for qty " + quantity);
  MinQty.Text = "Minimum Quantity: " + minQty;
  MinimumQty.Text = minQty.Replace("+", "");;
  sql = @"
      SELECT * FROM UnitPrices 
      WHERE ProductID=" + id + " ORDER BY description DESC";

	db = new DBHelper(sql);
  var colour = Util.Request("Decorations"); // Unbranded
  
  var subSQL = "Description LIKE '%" + Util.Request("Decorations") + @"%'
      AND Description NOT LIKE 'Additional%' AND (ISNULL(Timeframe, '') = '' OR Timeframe = '" +
      Util.Request("Timeframe") + "') ";
  //SQL.Text += subSQL; return;
  var attachDB = new DataTable();
  var attachDB2 = new DataTable();
  if (Util.Request("LanyardAttachment") != "")   attachDB = attachments.FilterTable("Attachment = '" + Util.Request("LanyardAttachment") + "'");
  if (Util.Request("LanyardAttachment2") != "") attachDB2 = attachments.FilterTable("Attachment = '" + Util.Request("LanyardAttachment2") + "'");

  var priceDB = db.FilterTable(subSQL);
  //SQL.Text += db.MakeHTMLTable();
  var addColorAmount = 0.0;
  var addColorAmount2 = 0.0;
  var addColorAmount3 = 0.0;
  if (colour != "" && colour != "Unbranded" && Util.Request("NumberColours") != "1") {
    var decoration = Util.Request("Decorations").ToLower();
    decoration = decoration.Replace("1 colour", "").Replace("1 col", "").Trim();
    //SQL.Text += "decoration = " + decoration + "<br>";
    var subsql = "Description LIKE 'Additional%' AND Description LIKE '%" + decoration + @"%' AND Timeframe = '" + Util.Request("Timeframe") + "'";
    //SQL.Text += subsql + "<br>"; // return;
    //subsql = "Description LIKE '%Additional Pad Print 4 x 4 cm%'";

    //SQL.Text += subsql; // return;
    var additionalDB = db.FilterTable("Description LIKE 'Additional%' AND Timeframe = '" + Util.Request("Timeframe") + "'");
    if (additionalDB.Rows.Count == 0) additionalDB = db.FilterTable(subsql);
    if (additionalDB.Rows.Count > 0) addColorAmount = (Util.ToInt(Util.Request("NumberColours"))-1) * Convert.ToDouble(additionalDB.Rows[priceRange]["UnitPrice"].ToString());
    if (additionalDB.Rows.Count > 0 && priceRange2 > 0) addColorAmount2 = (Util.ToInt(Util.Request("NumberColours"))-1) * Convert.ToDouble(additionalDB.Rows[priceRange2]["UnitPrice"].ToString());
    if (additionalDB.Rows.Count > 0 && priceRange3 > 0) addColorAmount3 = (Util.ToInt(Util.Request("NumberColours"))-1) * Convert.ToDouble(additionalDB.Rows[priceRange3]["UnitPrice"].ToString());
    //SQL.Text += "<p>addColorAmount= " + addColorAmount + " DB Count: " + additionalDB.Rows.Count;
  }
  //SQL.Text += "SELECT * FROM UnitPrices WHERE ProductID=" + id + " AND Description LIKE '%" + colour + @"%' AND (ISNULL(Timeframe, '') = '' OR Timeframe LIKE '%" + Util.Request("Timeframe") + "%') ";
  if (priceRange > priceDB.Rows.Count) Util.AlertAndEnd("No price found for " +  colour + " (column " + priceRange + ")");
  var unitPrice = 0.0;
  var unitPrice2 = 0.0;
  if (priceDB.Rows.Count > 0) {
    unitPrice = Convert.ToDouble(priceDB.Rows[priceRange]["UnitPrice"]) + addColorAmount;
    unitPrice2 = Convert.ToDouble(priceDB.Rows[priceRange2]["UnitPrice"]) + addColorAmount2;
  }
  var unitPrice3 = unitPrice2;

  //SQL.Text += "priceRange2 = " + priceRange2 + " priceRange3 = " + priceRange3 + DBHelper.MakeHTMLTable(priceDB);
  if (priceDB.Rows.Count > 0) unitPrice3 = Convert.ToDouble(priceDB.Rows[priceRange3]["UnitPrice"]) + addColorAmount3;
  
  
  //PriceColumn.Text += "Price $ = " + unitPrice.ToString();
      
      
 // If # of colors =2 THEN go read the "Additional" color
      
  //Debug.Text = sql;
  // Need to get the right price and * qty
  // if (priceDB.Rows.Count > 0) unitPrice = Convert.ToDouble(priceDB.Rows[0]["UnitPrice"]);

//unbranded|print|pad|wrap|full ?colour|colour|engrave text|engrave logo/text|engrave|heat transfer [0-9]{2} x [0-9]{2} cm|[1-2] sided
  foreach (DataRow row in db.Rows) {
    var decoration = row["Description"].ToString();
    //if (decoration.ToLower().Contains("1 col")) continue;
    if (decoration.ToLower().Contains("additional")) continue;
    decoration = decoration.Replace("Including", "").Replace("Additional", "");
    var selected = "";
    if (decoration.Contains( Util.Request("Decorations")) ) selected = " SELECTED";
    if (!Decorations.Text.Contains(decoration)) Decorations.Text += "<option" + selected + ">" + decoration + "</option>";
  }
  foreach (DataRow row in db.Rows) {
    var timeframe = row["TimeFrame"].ToString();
    var selected = "";
    if (timeframe == Util.Request("timeframe")) selected = " SELECTED";
    if (!TimeFrame.Text.Contains(timeframe)) TimeFrame.Text += "<option" + selected + ">" + timeframe + "</option>";
  }
  sql = "SELECT * FROM Packaging WHERE ProductId=" + id;
  //SQL.Text += sql; 
	db = new DBHelper(sql);
  double boxWeight = 0, qtyInBox = 0;
  double X = 10;
  double Y = 20;
  double Z = 20;
	if (db.Count > 0)
	{
		QtyPerBox.Text = db.Get("Quantity");
		WeightPerBox.Text = db.Get("GrossWeight");
    //Debug.Text += "width = " + db.Get("width");
		X = 20; if (Util.IsNumeric(db.Get("width")))  X = Util.ToDouble(db.Get("width"));
		Y = 20; if (Util.IsNumeric(db.Get("depth")))  Y = Util.ToDouble(db.Get("depth"));
		Z = 20; if (db.Get("height") == "") Z = Util.ToDouble(db.Get("height"));

    if (db.Get("GrossWeight") != "") boxWeight = Convert.ToDouble(db.Get("GrossWeight")) * Convert.ToDouble(db.Get("Quantity"));
    if (db.Get("Quantity") != "") qtyInBox = Convert.ToDouble(db.Get("Quantity"));
	}
  
  if (Util.Request("Quantity1") == "") return;  
  //Decoration.Text = "Colour";
  sql = "SELECT * FROM PrintArea WHERE ProductId=" + id;
	db = new DBHelper(sql);
  var s = db.Get("Dimensions"); // 40 mm (w) x 60 mm (h) possibly with ø for circulat ones
  if (s.Contains("ø"))
    PrintSize.Text = "Diameter " + Util.RightOf(s, "ø ");
  else
  {
    if (!s.Contains("mm"))
      PrintSize.Text = s;
    else
    {
      var width = Convert.ToDouble(Util.LeftOf(s, " mm (w)")) / 10.0;
      var length = Convert.ToDouble(Util.ParseBetween(s, "x ", " mm (h)")) / 10.0;
      PrintSize.Text = "Print Size: " + width + "cm x " + length + "cm";
    }
  }
  if (db.Get("location") != "") PrintSize.Text += " (" + db.Get("location") + ")";

  // Need to calc shipping based on the weight and number of boxes
  int boxCount = (int)Math.Ceiling(Convert.ToDouble(qty) / qtyInBox);
  var itemWeight = Math.Round(boxWeight / qtyInBox, 2); // grams to kg  

  // string, int, string, string, string, double, int, int, int)
  var nextWeekday = Util.NextWorkingDay(DateTime.Now);
  var shipping = 0.0;
  if (Util.Request("ShipCity") != "") shipping = GetShipping(nextWeekday.ToString("yyyy-MM-dd"), boxCount, Util.Request("ShipCity"), Util.Request("ShipState"), Util.Request("Postcode"), itemWeight, X, Y, Z);
  //ShippingAmount.Text = shipping.ToString("C");
  //Total.Text = (shipping + (unitPrice * quantity)).ToString("C");
  var setup = Convert.ToDouble(Util.Request("setupCost"));
  var numColors = Util.ToInt(Util.Request("NumberColours"));
  if (colour != "Unbranded") 
    setup *= Util.ToInt(Util.Request("NumberColours"));
  else
    setup = 0;
  
  QuoteString.Text = "<label>Quote Date:</label> " + DateTime.Now.ToString("dd-MMM-yy") + "<br>";
  QuoteString.Text += "<label>Product:</label> " + ProductCode.Text+ " - " + ProductName.Text + "<br>";
  
  var attachPrice = 0.0;
  var attachCode = Util.LeftOf( Util.Request("LanyardAttachment"), " -");
  var attachCode2 = Util.LeftOf( Util.Request("LanyardAttachment2"), " -");
  if (productCode.StartsWith("LAN"))
  {
    QuoteString.Text += "<label>Attachment:</label> " + Util.Request("LanyardAttachment");
    if (Util.Request("LanyardAttachment2") != "")  QuoteString.Text += " + " + Util.Request("LanyardAttachment2");

    if (attachDB.Rows.Count > 0) attachPrice = Convert.ToDouble(attachDB.Rows[0][priceRange+2].ToString());
    if (attachDB2.Rows.Count > 0) attachPrice += Convert.ToDouble(attachDB2.Rows[0][priceRange+2].ToString());
    QuoteString.Text += " - " + attachPrice.ToString("C") + " each<br>";
    
  }
  QuoteString.Text += GetQuoteString(setup, shipping, quantity, unitPrice, numColors, addColorAmount, attachPrice);
  
  // ** Now see if they want a second quote for a different qty
  if (quantity2 > 0) {
    QuoteString.Text += "<hr><small><b>Second Quote</b></small><br>";
    boxCount = (int)Math.Ceiling(quantity2 / qtyInBox);
    if (Util.Request("ShipCity") != "") shipping = GetShipping(nextWeekday.ToString("yyyy-MM-dd"), boxCount, Util.Request("ShipCity"), Util.Request("ShipState"), Util.Request("Postcode"), itemWeight, X, Y, Z);
    QuoteString.Text += GetQuoteString(setup, shipping, Util.ToInt(quantity2), unitPrice2, numColors, addColorAmount2, attachPrice);
  }
  
  // ** Now see if they want a 3rd quote for a different qty
  if (quantity3 > 0) {
    QuoteString.Text += "<hr><small><b>Third Quote</b></small><br>";
    boxCount = (int)Math.Ceiling(quantity3 / qtyInBox);
    if (Util.Request("ShipCity") != "") shipping = GetShipping(nextWeekday.ToString("yyyy-MM-dd"), boxCount, Util.Request("ShipCity"), Util.Request("ShipState"), Util.Request("Postcode"), itemWeight, X, Y, Z);
    
    QuoteString.Text += GetQuoteString(setup, shipping, Util.ToInt(quantity3), unitPrice3, numColors, addColorAmount3, attachPrice);
  }
  
  if (Util.Request("Email") != "") {
    var msg = QuoteString.Text + "<img src='/images/products/detail/" + ProductCode.Text + ".jpg'><br>";
    var subj = "The Resource Group - Pricing for " + ProductName.Text;
    Util.SendMail(Util.Request("Email"), "sales@resgr.com.au", subj, msg);
  
  }
}

string GetQuoteString(double setup, double shipping, int quantity, double unitPrice, int colors, double additionalPerColor, double lanyardAttach)
{
  var total = (shipping + setup + (unitPrice * quantity));
  var quote = "";
  if (lanyardAttach > 0)  
    quote  += "<label>Attachment $:</label> " + (quantity * lanyardAttach).ToString("C") + "<br>";
   
  if (colors > 1) quote += "<label>Additional Colors:</label> " + additionalPerColor.ToString("C") + " (per color) * " + quantity.ToString() + " = " + (quantity * (colors-1) * additionalPerColor).ToString("C") + "<br>";

  quote += "<label> Price:</label> " + (unitPrice * quantity).ToString("C") + " (Qty: " + quantity.ToString() + " at " + unitPrice.ToString("C") + ")<br>";
  quote += "<label>Setup:</label> " + setup.ToString("C") + " (" + colors + " Colours)<br>";
  if (shipping == 0)
    quote += "<label>Shipping:</label> <span style='color: maroon;'>For a quote with freight please enter city, state & postcode above</span>"; 
  else
    quote += "<label>Shipping:</label> " + shipping.ToString("C") + "<br>";
    
  quote += "<label>TOTAL:</label> <b>" + total.ToString("C") + "</b><br>";
  if (lanyardAttach > 0)  
    quote  += "<label>GRAND TOTAL:</label> <b>" + (total + (quantity * lanyardAttach)).ToString("C") + "</b> (Prices Exclude GST)<br>";
  return quote + "<br>";
}

double GetShipping(string pickupDate, int boxCount, string city, string state, string postalCode, double weight, double X, double Y, double Z)
{
  var url = "https://www.tntexpress.com.au/Rtt/inputRequest.asp";
  var xml = @"<?xml version='1.0'?>
<enquiry xmlns='http://www.tntexpress.com.au'>
<ratedTransitTimeEnquiry>
  <cutOffTimeEnquiry>
    <collectionAddress>
      <suburb>Glendenning</suburb>
      <postCode>2761</postCode>
      <state>NSW</state>
    </collectionAddress>
    <deliveryAddress>
      <suburb>" + city + @"</suburb>
      <postCode>" + postalCode + @"</postCode>
      <state>" + state + @"</state>
    </deliveryAddress>
    <shippingDate>" + pickupDate + @"</shippingDate>
    <userCurrentLocalDateTime>
      2017-09-13T10:00:00
    </userCurrentLocalDateTime>
    <dangerousGoods>
      <dangerous>false</dangerous>
    </dangerousGoods>
    <packageLines packageType=""D"">
      <packageLine>
        <numberOfPackages>" + boxCount + @"</numberOfPackages>
        <dimensions unit=""cm"">
          <length>" + Util.ToInt(Y) + @"</length>
          <width>" + Util.ToInt(X) + @"</width>
          <height>" + Util.ToInt(Z) + @"</height>
        </dimensions>
        <weight unit=""kg"">
          <weight>" + weight + @"</weight>
        </weight>
      </packageLine>
    </packageLines>
  </cutOffTimeEnquiry>
  <termsOfPayment>
    <senderAccount>30021588</senderAccount>
    <payer>S</payer>
  </termsOfPayment>
</ratedTransitTimeEnquiry>
</enquiry>";
  var results = Util.PostURL(url, "username=CIT00000000000101392&password=prodTNT123&version=2&xmlRequest=" +  HttpUtility.UrlEncode(xml) );

  if (results.Contains("brokenRules")) Debug.Text += Util.ParseBetween(results, "<description>", "</description>");
Debug.Text = xml + "\n\nRESULTS\n\n" + results;

  var price = Util.ParseBetween(results, "<price>", "</price>");
  if (price == "") price = "0";
  var shipAmount = Math.Round(Convert.ToDouble(price) * 1.20d, 2);

  return shipAmount;
}
</script>