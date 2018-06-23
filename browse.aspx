<%@ Page Language="C#" AutoEventWireup="false" debug="true" %>
    <%@ Import Namespace="System.IO" %>

        <!--#include file="headerNEW.aspx"-->

        <link rel="stylesheet" href="css/browse.css">

        <!--#include file="leftNav.aspx"-->

        <div id="main" class="grid_8">
            <%
    var mid = Util.Request("mid");
    if (mid == "") return;
    if (!Util.IsInt(mid)) Util.AlertAndEnd("MID is not integer");
    var sql = @"SELECT * FROM Categories WHERE catID=" + mid;
    db = new DBHelper(sql);
    if (db.Count == 0) Util.AlertAndEnd("Invalid Category ID " + mid);
    var category = db.Get("catName");

    sql = "select categories.catid,catname from CategoryRelations inner join categories on CategoryRelations.catid=categories.catid where incatid=" + mid + " ORDER BY catname";
    DBHelper subCats = new DBHelper(sql);
    var sort = "ProductName";
    if ( mid == "171") sort = "case WHEN SUBSTRING(ProductCode, 1, 3) = 'LAN' then 'a' else 'z' END, catname, case WHEN IsNumeric(SUBSTRING(ProductCode, 4, 3)) = '1' then Convert(int, SUBSTRING(ProductCode, 4, 3)) else 99999 END, ProductCode";
 
    if (mid == "284" || mid == "171")
        sql = @"select top 200 catname, Product.* from ProductInSubCategory INNER JOIN Product ON Product.ProductID=ProductInSubCategory.ProductID 
        INNER JOIN Categories On Categories.CatID = ProductInSubCategory.inSubCatID
        WHERE insubcatid=" + mid + " or insubcatid IN (SELECT catID FROM CategoryRelations WHERE inCatID=" + mid + ") ORDER BY " + sort;
    else
        sql = "select top 200 Product.*, '' as catname from ProductInSubCategory INNER JOIN Product ON Product.ProductID=ProductInSubCategory.ProductID where insubcatid=" + mid + " or insubcatid IN (select catID FROM CategoryRelations WHERE inCatID=" + mid + 
            ") ORDER BY ProductCode";
    //Response.Write(sql);
    DBHelper productsDB = new DBHelper(sql); // ProductCode ProductName ProductDescription

    // EXAMPLE of how to over-ride default image by a category number
    if (mid == "284") { %>
                <br /><img src="/images/headers/CROSS PENS.jpg" alt="Cross Pens" "/><br /><br /><br />
        <p style="text-align:center; ">Cross luxury writing instruments are renowned for their distinctive style and superior quality. They have long been symbols of success and achievement, and no longer serve as simply functional writing instruments. </p>
        <p style="text-align:center; ">Cross Pens was Founded in 1846 by Richard Cross, when craftsman, inventor and son Alonzo Townsend Cross began to design decorative pencil cases in gold, silver and brass.  Since that time five generations of the Cross family have refined the jeweller’s craft with exacting design and precise tooling to produce the first high quality writing instrument.  For over 165 years A.T. Cross Company has reigned as the premiere manufacturer of quality writing instruments and is now sold in over 150 countries around the world.</p> 
        <p style="text-align:center; ">Unlike many pen manufacturers, Cross Pens come with a lifetime mechanical warranty.</p>
    <% } else if (mid == "171 ") { // Lanyards Header %>
        <img src="/images/headers/Lanyards.jpg " alt="Lanyards " style="border:1px solid #dadada; margin-left:5px; "><br /><br />
        <h3>Popular Lanyards</h3>
        <div style="border:1px solid #dadada;width:366px;height:250px;margin:1px 0px 5px 5px;padding:0px;display:block;float:left; ">
        <span class="tnc ">LAN3</span>
        <div class="tnidiv ">
            <a href="/product.aspx?pid=401 "><img src="/images/LAN3top.jpg " alt="Boot Lace Style (tubular) "></a>
        </div>
        <span class="tnn "><a href="/product.aspx?pid=401 " style="font-size:1.3em; ">Printed Bootlace Style Lanyard (tubular)</a></span>
        </div>
        <div style="border:1px solid #dadada;width:366px;height:250px;margin:1px 0px 5px 5px;padding:0px;display:block;float:left; ">
        <span class="tnc ">LAN4</span>
        <div class="tnidiv ">
            <a href="/product.aspx?pid=402 ">
                <img src="/images/LAN4top.jpg " alt="Polyester Ribbed Lanyard (flat). ">
            </a>
        </div>
        <span class="tnn ">
            <a href="/product.aspx?pid=402 " style="font-size:1.3em; ">Printed Polyester Ribbed Lanyard (flat)</a>
        </span>
        </div> 
        
        <% } else if (mid == "281 ") { // Budget Pens %>
        <br /><img src="/images/headers/Budget Pens.jpg " alt="Budget Pens " "/><br /><br />
                <p>Budget Pens is the new pen range brought to you by Specialty Range. Budget Pens is a very competitive pen range with over 100 pens and colours to choose from. All budget pens are held in our warehouse overseas where we print and freight
                    pens based on demand. Turnaround time is strictly 2-3 weeks from order confirmation. Orders can be done faster but additional freight charges may apply.</p>
                <p> If your client is after 2, 3 or 4 colour print on their pen they might want to consider our budget pen selection. Decoration gets cheaper as the colours and complexity increases.</p>

                <% } else if (mid == "000") { // Michelle change 000 to category  %>
                    <br /><img src="/images/headers/XXXXXX"><br /><br />
                    <p>Some text </p>

                    <% } else // Normal Category
            Response.Write("<img src='/images/headers/" + category + ".jpg' alt='" + category + "' class='product-header-image' />");

// DOES This have Sub-Categories? If so display each as a link
if (subCats.Rows.Count > 0 && mid != "284" && mid != "171") // Lanyards Header)
{
    Response.Write("<h5>Sub-Categories</h5><ul>");
    foreach (DataRow row in subCats.Rows)
        Response.Write("<li style='text-align:left'><a href=?mid=" + row["catID"] + ">" + row["catName"] + "</a></li>");
    Response.Write("</ul>");
}
Response.Write("<div class='tnc_container'>");

// Loop over each product in category
var lastCategory = "";
foreach (DataRow row in productsDB.Rows)
{
    var cat = row["catName"].ToString();
    var name = row["ProductName"].ToString();
    if (mid == "284" || mid == "171")
    {
        if (lastCategory != cat) {
            Response.Write("<br clear=left><h3> " + cat + " </h3>");
            lastCategory = cat;
        }
    }
    %>
                        <div id="tn">
                            <span class="tnc"><%= row["ProductCode"] %></span>
                            <div class="tnidiv">
                                <a href='/product.aspx?pid=<%= row["productId"] %>'><img src='/images/products/thumbnail/<%= row["productCode"] %>.jpg' alt='<%= row["ProductDescription"] %>' /></a>
                            </div>
                            <% if (name.Length > 45) { %>
                                <span class="tnn"><a href='/product.aspx?pid=<%= row["productId"] %>' style="font: 13px arial;"><%= name %></a></span>
                                <% } else { %>
                                    <span class="tnn"><a href='/product.aspx?pid=<%= row["productId"] %>'><%= name %></a></span>
                                    <% } %>
                        </div>
                        <% } %>
        </div>
        </div>
        <!--#include file="footerNEW.aspx"-->