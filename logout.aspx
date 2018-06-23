<%@ PAGE language="C#" AutoEventWireup="false" EnableSessionState="true" %>
<%
	System.Util obj = new System.Util();
	Session["user"] = "";
	Response.Expires = -100; // Tell Browser not to cache logout HTML (really JavaScript)
%>
<meta http-equiv="Expires" content="NONE">
<script>alert("Logging out of the system..."); document.location = "/";</script>

