<%@ Page Language="C#" AutoEventWireup="false"  %>

    <!--#include file="headerNEW.aspx"-->

    <link rel="stylesheet" href="css/home.css">

    <%
    db = new DBHelper("select top 200 * from Categories where cattype='M' order by DisplayOrder, catname"); 
    string sideColor = "#fff";
    if (serverName.Contains("custx")) sideColor = "green";
%>

        <div class="container content">
            <%
	foreach (DataRow row in db.Rows)
	{
		if( Convert.ToInt32( row["displayOrder"] ) > 0 )
		{
			Response.Write("<div class='frontpage-category-photos-wrapper'>");
			Response.Write("<div class='frontpage-category-photos'>");
			Response.Write( "<a href=/browse.aspx?mid=" + row["catId"] + ">");
			Response.Write(     "<div class='category-img-wrapper'><img src='/images/categories/" + row["catName"] + ".jpg' /></div>");
			Response.Write(     "<div class='category-name'>" + row["catName"] + "</div>");
			Response.Write( "</a>");
			Response.Write("</div>");
			Response.Write("</div>");
		}
	}
%>
        </div>
        <!--#include file="footerNEW.aspx"-->