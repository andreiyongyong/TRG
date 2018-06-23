<%@ Page Language="C#" AutoEventWireup="false" %>
    <!--#include file="headerNEW.aspx"-->
    <!--#include file="leftNav.aspx"-->
    <%
	string sql = "SELECT * FROM Users WHERE approved=1 AND username='" + Util.EncodeRequest("txtusername") + "' AND UserPass='" + Util.EncodeRequest("txtpassword") + "'";
//'  Response.Write(sql); return;

	db = new DBHelper(sql);
	if (db.Count == 0) Util.AlertAndEnd("Invalid Username or Password.");
	
	Util.SetCookie("ckusername", Util.Request("txtusername"), 90);
	Session["user"] = Util.EncodeRequest("txtusername");

	if (Request.ServerVariables["HTTP_REFERER"].Contains(".aspx"))
		Response.Write("<script>document.location='" + Request.ServerVariables["HTTP_REFERER"] + "';</script>");
	else
		Response.Write("<script>document.location='/';</script>");

%>