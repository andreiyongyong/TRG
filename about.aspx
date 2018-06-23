<%@ Page Language="C#" AutoEventWireup="false" %>
    <!--#include file="headerNEW.aspx"-->
    <!--#include file="leftNav.aspx"-->
    <div id="main" class="grid_8">
        <h2>About Us</h2>
        <h4>The Resource Group is an Australian owned and run business, selling quality promotional products to over 1,700 distributors nation-wide.</h4>
        <br />
        <h3 style="border-bottom:1px solid #ebebeb;">Why become a TRG distributor?</h3>

        <ul>
            <li>Huge <b>RANGE</b> of quality products</li>
            <li>Large amounts of <b>LOCAL STOCK</b></li>
            <li><b>FAST</b> turnaround – 2 weeks for local products</li>
            <li><b>QUALITY</b> print / print control in house</li>
            <li><b>VARIETY</b> of brand applications – screen-print, pad-print, engrave etc</li>
            <li><b>NO MERCHANT FEES</b> on Visa and Mastercard. We understand merchant charges are our cost of doing business</li>
            <li>30 day <b>ACCOUNTS</b> – giving you more flexibility</li>
            <li>We only sell to <b>APPROVED</b> promotional distributors.</li>
            <li>We endeavour to source specific products to <b>MEET YOUR DEMANDS</b>.</li>
            <li style="list-style-type: none;">&nbsp; </li>
            <li style="list-style-type: none;">
                <p style="font:bold 1.3em Georgia, 'Times New Roman', Times, serif;color:#444;padding:0px;">Become a TRG Distributor today!</p>
            </li>
        </ul>
        <br />

        <% // Response.Write(GetInventory("172")); %>
            <!--script src="http://www.hitpromo.net/product/widget?number=172&nocss=0" type="text/javascript"></script-->
            <div id="product-container"></div>
            <% if (Util.GetSession("user") == "") return; %>

                <h3 style="border-bottom:1px solid #ebebeb;">Terms and Conditions of Sale</h3>
                <ul>
                    <li><a href="TRG_Terms_&_Conditions.pdf" style="font:bold 1.3em Georgia, 'Times New Roman', Times, serif;text-decoration: none;" onmouseover="this.style.color='#CA171D'" onmouseout="this.style.color='#444'">Click here to view our Terms & Conditions of Sale.</a></li>
                </ul><br>
                <h5>Contact Details</h5>
                <table class="about_table" style="border:0px">
                    <tbody>
                        <tr>
                            <td class="align-right" width="90px">Telephone</td>
                            <td class="align-left">(02) 8882 1888</td>
                        </tr>
                        <tr>
                            <td class="align-right">Fax</td>
                            <td class="align-left">(02) 9677 0554</td>
                        </tr>
                        <tr>
                            <td class="align-right">Address</td>
                            <td class="align-left">1/7 Kelham Pl<br />Glendenning<br />NSW 2761</td>
                        </tr>
                        <tr>
                            <td class="align-right">Email</td>
                            <td class="align-left">production@resgr.com.au<br />sales@resgr.com.au</td>
                        </tr>
                    </tbody>
                </table>
    </div>
    <!--#include file="footerNEW.aspx"-->
    <script runat="server">
        string GetInventory(string productNo) {
            var util = new System.Util();

            byte[] bPostArray = Encoding.UTF8.GetBytes("username=theresourcegroupptyltd&password=trg123");
            System.Net.WebClient wc = new System.Net.WebClient();

            byte[] bHTML = wc.UploadData("https://www.hitpromo.net/site/login", "POST", bPostArray);
            System.Net.WebHeaderCollection myHeaders = wc.ResponseHeaders;

            var myCookie = myHeaders["Set-Cookie"];
            myCookie += "; 758a58fe75ac0e86b278a2cffd65cc9f=788596e49df6fe10093f9b8921592fd2f9b88971a%3A4%3A%7Bi%3A0%3Bs%3A22%3A%22theresourcegroupptyltd%22%3Bi%3A1%3Bs%3A22%3A%22theresourcegroupptyltd%22%3Bi%3A2%3Bi%3A1209600%3Bi%3A3%3Ba%3A5%3A%7Bs%3A11%3A%22customer_id%22%3Bs%3A6%3A%22107676%22%3Bs%3A15%3A%22customer_number%22%3Bs%3A6%3A%22779946%22%3Bs%3A11%3A%22hitEmployee%22%3Bb%3A0%3Bs%3A7%3A%22subUser%22%3Bb%3A0%3Bs%3A19%3A%22allowChangePassword%22%3Bb%3A1%3B%7D%7D";
            wc.Headers.Add(System.Net.HttpRequestHeader.Cookie, myCookie);

            var wholePage = new UTF8Encoding().GetString(wc.DownloadData("https://www.hitpromo.net/product/show/172/"));
            var justTable = Util.ParseBetween(wholePage, "<h2>Inventory</h2>", "</div>");
            var startTable = Util.LeftOf(justTable, "<tbody>");
            var tableBody = Util.ParseBetween(justTable, "<tbody>", "</tbody>");
            var returnString = startTable + "<tbody>";
            foreach(string row in tableBody.Split(new string[] {
                "<tr>"
            }, StringSplitOptions.None)) {
                var name = Util.ParseBetween(row, "<td>", "</td>");
                if (name == "") continue;
                var count = Util.ParseBetween(row, "</td>\n<td>", "</td>");
                var num = Convert.ToInt32(count.Replace(",", ""));
                if (num > 2500) count = "2,500+";
                returnString += "<tr><td>" + name + "</td><td> " + count + "</td></tr>\n";
            }
            return returnString + "</tbody></table>";
        }
    </script>