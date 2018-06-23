<%@ Page Language="C#" AutoEventWireup="false" %>
    <!--#include file="headerNEW.aspx"-->
    <!--#include file="leftNav.aspx"-->

    <div id="main" class="grid_8">
        <h2>Pricelist & Catalogue</h2>
        <br /><br />
        <% if (Util.GetSession("user") != "") { %>
            <div style="display: block; margin-left: auto; margin-right: auto; height: 310px;"><br />

                <a href="TRG_SpecialtyRange_PriceList.xlsx">
                    <div style="float:left;text-align: center;width:180px;font:bold 1em/2.4em Georgia, 'Times New Roman', Times, serif;color:#444;font-size:1.5em;font-weight:normal;letter-spacing:-0.5px;padding:3px">
                        <img alt="Download Our Latest Pricelist" src="/images/pricelist_excel.jpg" style="display: block;margin-left: auto;margin-right: auto; height:170px;" />
                        Latest Pricelist (Updated) 
                        <img alt="" src="images/download.jpg" style="width:100px;" />
                    </div>
                </a>

                <div style="float:left;text-align: center;width:180px;font:bold 1em/2.4em Georgia, 'Times New Roman', Times, serif;color:#444;font-size:1.5em;font-weight:normal;letter-spacing:-0.5px;padding:3px">
                    <img alt="Download Our Latest Catalogue" src="/images/catalogue_2014.jpg" style="display: block;margin-left: auto;margin-right: auto; height:170px;" />
                    2017 Catalogue (coming soon) 
                    <img alt="" src="images/download.jpg" style="width:100px;" />
                </div>

                <div style="margin-left: 15px; float:left;text-align: center;width:180px; height: 225px; border:0px solid skyblue;">
                    <b>
                        <a href="export.aspx" style="color: blue; text-decoration: underline;"> Export Products<br />
                            <br />
                            <img alt="" src="http://4.bp.blogspot.com/-kjCDQoF-paY/WQHjVbgON8I/AAAAAAAADfY/9vS1eBjN8rQZ8sXrWVIaSVJTAZMbH1HTQCLcB/s1600/images.jpg" height="170" />
 
                            <img alt="" src="images/download.jpg" style="width:100px; margin-top: 40px;" />
                        </a>
                    </b>

                </div>
            </div>
            <br />
    </div>
    <% } else { %>
        Please login to access the current catalogue and pricelist.
        <% } %>

</div>
    <!--#include file="footerNEW.aspx"-->