<%
	db = new DBHelper("select top 200 * from Categories where cattype='M' order by catname"); 
	string sideColor = "#fff";
	if (serverName.Contains("custx")) sideColor = "green";
%>

    <div id="left-columns" class="grid_4">
        <% 
    if (Util.GetSession("user") == "") { %>
            <form class="input-form" action="/login.aspx" method="post" id="loginform" style="border:1px solid #dadada;width:156px;height:118px;background-color:#fff;margin-left:2px;margin-top:3px;margin-bottom:0px;">
                <span class="float-left" style="color:#3d3d3d;border-bottom:1px dashed #ddd;">Distributor Login</span>
                <br />
                <input id="txtusername" class="input" type="text" name="txtusername" value="<%= Util.GetCookie(" ckusername ") %>" onfocus="this.value='';this.style.fontStyle='normal';" style="width:144px;font-style:italic;background:#FFF;z-index:999999;" />
                <input id="txtpassword" class="input" type="password" name="txtpassword" value="" style="width:144px;background:#FFF;z-index:999999;" />
                <input id="btn_login" type="submit" value="Login" class="input button float-right" style="margin-top: 5px; margin-right: 2px; height: 25px;" />
            </form>
            <% } else { %>
                <div style="border:1px solid #dadada;width:170px;height:50px;background-color:#fff;margin: 3px 0 5px 3px;padding:0px;">
                    <div style="text-align:center;color:#3d3d3d;margin-top:8px;">
                        Welcome
                        <%= Util.GetSession("user") %><br>
                            <a href="/logout.aspx">LOGOUT</a>
                    </div>
                </div>
                <div style="margin: 15px 0 20px 45px;">
                    <b><a href="export.aspx" style="color: blue; text-decoration: underline;"> Export Products</a></b><br>
                </div>
                <% } %>

                    <!-- BEGIN: Constant Contact Email List Form Button -->
                    <div align="center"><a href="https://visitor.r20.constantcontact.com/d.jsp?llr=coqvjwiab&amp;p=oi&amp;m=1108849902018&amp;sit=omroczpgb&amp;f=3ca2542d-d86b-4841-91d5-1a2d842675e2" class="button" style="background-color: rgb(0, 0, 0); border: 1px solid rgb(91, 91, 91); color: rgb(255, 255, 255); display: inline-block; padding: 8px 10px; text-shadow: none; border-radius: 20px;">NEW PRODUCTS and SPECIALS</a>
                        <!-- BEGIN: Email Marketing you can trust -->
                        <div id="ctct_button_footer" align="center" style="font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#999999;margin-top: 10px;">For Email Marketing you can trust.</div>
                    </div>

                    <div class="sidemenu">
                        <table style="padding:0px;width:178px;margin-top:5px;">
                            <tr>
                                <td style="background-color:<%= sideColor %>">
                                    <ul>
                                        <%
                        foreach (DataRow row in db.Rows)
                            Response.Write("<li><a href=/browse.aspx?mid=" + row["catID"] + ">" + row["catName"] + "</a></li>");
                        %>
                                    </ul>
                                    <br />
                                </td>
                            </tr>
                        </table>
                    </div>
    </div>