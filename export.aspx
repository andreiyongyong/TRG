<%@ PAGE Language="C#" Debug="true" %>
<% string title = "Export"; %>
<!--#include file="headerNEW.aspx"-->


<script runat="server">
void Page_Load()
{
  var dsn = HttpContext.Current.ApplicationInstance.Application["DSN"].ToString();
  var sql = @"
    SELECT ProductCode, ProductName, Width, Depth, Height, Diameter
      , SUBSTRING(Colours, 1, 50) as Colors
      , PrintArea=(SELECT TOP 1 Location + ' ' + Dimensions FROM PrintArea WHERE ProductId=Product.ProductId)
      , SetupCost=ISNULL(SetupCost, 50)
      , Thumbnail='images/products/thumbnail/' + ProductCode + '.jpg'
      , Image='images/products/hires/' + ProductCode + '.jpg'
      , ProductDescription
    FROM Product 
    LEFT OUTER JOIN Colours ON Colours.ProductID=Product.ProductID
    LEFT OUTER JOIN Dimensions ON Dimensions.productID=Product.ProductID
    ORDER BY ProductCode";
  var db = new DBHelper(sql, dsn);
  //Response.Write( db.MakeHTMLTable() );
  db.StreamExcel("ProductExport.xlsx"); // This will stop here (Response.End)
}
</script>