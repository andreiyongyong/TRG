<%@ Page Language="C#" AutoEventWireup="false" %>
    <!--#include file="headerNEW.aspx"-->

    <link rel="stylesheet" href="css/style.css">
    <link rel="stylesheet" href="css/productSuggestions.css">

<%
    // variables to be used later 
    var isBudgetPen = false;
    var lastDesc = "";
    var lastTimeframe = "";
    int rowNum = 0;
    int nPricingHeaderColumns = 0;

    var pid = Util.Request("pid");
    if (pid == "" || !Util.IsNumeric(pid)) return;
    var sql = @"SELECT * FROM Product 
      LEFT OUTER JOIN Colours ON Colours.ProductID=Product.ProductID 
      LEFT OUTER JOIN Dimensions ON Dimensions.ProductID=Product.ProductID 
      LEFT OUTER JOIN Packaging  ON Packaging.ProductID=Product.ProductID 
      LEFT OUTER JOIN AdditionalContent ON AdditionalContent.ProductID=Product.ProductID 
      WHERE Product.productID=" + pid;
    db = new DBHelper(sql);
    if (db.Rows.Count == 0) Util.AlertAndEnd("Invalid Product ID " + pid);
    var productCode = db.Get("productCode");
    var prefixProductCodeLength = ( productCode.Length < 3 )? productCode.Length : 3; 
    var prefixProductCode = productCode.Substring(0, prefixProductCodeLength);
    var FreightCodeId = db.Get("FreightCodeId");
    var digitalSetup = db.Get("DigitalSetup");
    var productSetupCost = db.Get("SetupCost");
    var productRepeatSetup = db.Get("RepeatSetup");
    var isFullColourPrint = (db.Get("IsFullColourPrint").ToString() == "1");
    var isScreenPrint = (db.Get("IsScreenPrint").ToString() == "1");
    var isPadPrint = (db.Get("IsPadPrint").ToString() == "1");
    var isWrapPrint = (db.Get("IsWrapPrint").ToString() == "1");
    var codes = new string[20]; //' array of product codes
    codes[0] = "one";
    for (int i=1; i < 20; i++) codes[i] = db.Get("prodCode" + i);

    System.Globalization.TextInfo textInfo = new System.Globalization.CultureInfo("en-US",false).TextInfo;  // Creates a TextInfo based on the "en-US" culture.
%>
    <script runat="server">
        // -----------------------------
        public int getPricingDisplayRankingOrder(string timeframe,
                string description) {
                //        Regex regNumber = new Regex(@"[0-9]+");

                int returnValue = 0;

                if (timeframe.ToLower().Contains("month")) {
                    returnValue = 6000 + getPricingDisplayRankingOrderFromNumber(timeframe) + getPricingDisplayRankingOrderFromDescription(description);;
                } else if (timeframe.ToLower().Contains("days rush")) {
                    returnValue = 5000 + getPricingDisplayRankingOrderFromNumber(timeframe) + getPricingDisplayRankingOrderFromDescription(description);;
                } else if (timeframe.ToLower().Contains("week rush")) {
                    returnValue = 4000 + getPricingDisplayRankingOrderFromNumber(timeframe) + getPricingDisplayRankingOrderFromDescription(description);;
                } else if (timeframe.ToLower().Contains("week")) {
                    returnValue = 3000 + getPricingDisplayRankingOrderFromNumber(timeframe) + getPricingDisplayRankingOrderFromDescription(description);;
                } else if (timeframe.ToLower().Contains("same day")) {
                    returnValue = 1000 + getPricingDisplayRankingOrderFromNumber(timeframe) + getPricingDisplayRankingOrderFromDescription(description);
                } else if (timeframe.ToLower().Contains("day")) {
                    returnValue = 2000 + getPricingDisplayRankingOrderFromNumber(timeframe) + getPricingDisplayRankingOrderFromDescription(description);;
                }

                return returnValue;

            } // getPricingDisplayRankingOrder

        public int getPricingDisplayRankingOrderFromNumber(string timeframe) {
                int returnValue = 0;
                Regex re = new Regex(@"\d+");
                Match m = re.Match(timeframe);
                if (m.Success) {
                    returnValue = Int32.Parse(m.Value) * 100;
                }
                return returnValue;

            } // getPricingDisplayRankingOrderFromNumber 

        public int getPricingDisplayRankingOrderFromDescription(string description) {
                int returnValue = 0;

                if (description.ToLower().Contains("pad")) {
                    returnValue += 10;
                } else if (description.ToLower().Contains("wrap")) {
                    returnValue += 20;
                } else if (description.ToLower().Contains("full")) {
                    returnValue += 30;
                } else if (description.ToLower().Contains("engrave")) {
                    returnValue += 40;
                } else if (description.ToLower().Contains("heat transfer")) {
                    returnValue += 50;
                }

                if (description.ToLower().Contains("including")) {
                    returnValue += 1;
                } else if (description.ToLower().Contains("additional")) {
                    returnValue += 2;
                }

                return returnValue;

            } // getPricingDisplayRankingOrderFromDescription 


        public void writeOutLanyardClipOptionsForPriceCalculator() {
                Response.Write("dialogHTML += \"<option value='A1' data-image='/images/A1.jpg' data-title='A1' data-description='Simple J Hook'>A1</option>\";");
                Response.Write("dialogHTML += \"<option value='A2' data-image='/images/A2.jpg' data-title='A2' data-description='Metal Alligator Clip'>A2</option>\";");
                Response.Write("dialogHTML += \"<option value='A3' data-image='/images/A3.jpg' data-title='A3' data-description='Plastic Alligator Clip'>A3</option>\";");
                Response.Write("dialogHTML += \"<option value='A3A' data-image='/images/A3A.jpg' data-title='A3A' data-description='Plastic Simple J Hook'>A3A</option>\";");
                Response.Write("dialogHTML += \"<option value='A4' data-image='/images/A4.jpg' data-title='A4' data-description='Crab claw'>A4</option>\";");
                Response.Write("dialogHTML += \"<option value='A5' data-image='/images/A5.jpg' data-title='A5' data-description='Deluxe Swivel'>A5</option>\";");
                Response.Write("dialogHTML += \"<option value='A6' data-image='/images/A6.jpg' data-title='A6' data-description='Executive Swivel'>A6</option>\";");
                Response.Write("dialogHTML += \"<option value='A7' data-image='/images/A7.jpg' data-title='A7' data-description='J Swivel'>A7</option>\";");
                Response.Write("dialogHTML += \"<option value='A8' data-image='/images/A8.jpg' data-title='A8' data-description='Superior O Clip'>A8</option>\";");
                Response.Write("dialogHTML += \"<option value='01' data-image='/images/01.jpg' data-title='01' data-description='Safety Breakaway'>01</option>\";");
                Response.Write("dialogHTML += \"<option value='B2' data-image='/images/B2.jpg' data-title='B2' data-description='o Ring'>B2</option>\";");
                Response.Write("dialogHTML += \"<option value='B3' data-image='/images/B3.jpg' data-title='B3' data-description='Plastic Snap Loop'>B3</option>\";");
                Response.Write("dialogHTML += \"<option value='PH1' data-image='/images/PH1.jpg' data-title='PH1' data-description='Detachable Phone Holder'>PH1</option>\";");
                Response.Write("dialogHTML += \"<option value='PH2' data-image='/images/PH2.jpg' data-title='PH2' data-description='Phone Holder'>PH2</option>\";");

            } // writeOutLanyardClipOptionsForPriceCalculator
    </script>

        <div id="main" class="">
            <div class="content-section">
                <div class="container">
                    <div class="row">
                        <div class="route">
                            <div class="notification">
                                <span>Home </span>
                                <span class="glyphicon glyphicon-arrow-right"></span>
                                <span>All Products </span>
                                <span class="glyphicon glyphicon-arrow-right"></span>
                                <span>Drinkware</span>
                                <span class="glyphicon glyphicon-arrow-right"></span>
                                <span style="color:#11c1ee"><%= db.Get("ProductCode") %> <%= db.Get("ProductName") %></span>
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-md-6">
                            <div class="row">
                                <div class="col-md-2">
                                    <div class="product-item-1">
    <% 
        var path = Server.MapPath("images/products/thumbnail/");
        var productThumbnails = 0;
        for (int i = 2; i < 20; i++){
            if (File.Exists(path + productCode + "-" + i + ".jpg")) {
                productThumbnails = i;
            } 
        }
                                        
    %>
                                        <div class="c_left_vertical_arrow_item">
    <%
        if (productThumbnails > 4) {
    %>
                                                <span class="glyphicon glyphicon-triangle-top upper-triangle"></span>
    <%  
        } 
    %>
                                        </div>

    <%
        if (productThumbnails > 0) {
            for (int i=2; i <= productThumbnails ; i++)
            {
    %>
                                        <a href="?pid=<%= pid %>&code=<%= codes[i] %>&no=<%= i %>">
                                            <div class="product-thumb c_left_vertical_item" no="<%= i %>">
                                                <img src="/images/products/thumbnail/<%= productCode %>-<%= i %>.jpg" alt="" OnMouseover="javascript:swap_image('productimage','','images/products/detail/<%= productCode %>-<%= i %>.jpg',1);fnc_hires('hireslink','<%= productCode %>-<%= i %>')">
                                            </div>
                                        </a>
    <% 
            }
        }
    %>
                                        <!--<div class="product-thumb selected">
                                            <img src="images/products/B702-14_rec.png" alt="">
                                        </div>-->
                                        <!-- /.product-thumb -->

                                        <div class="c_left_vertical_arrow_item">
    <%
        if (productThumbnails > 4) {
    %>
                                                <span class="glyphicon glyphicon-triangle-bottom bottom-triangle"></span>
    <% 
        } 
    %>
                                        </div>

                                        <div class="c_left_vertical_new_div">new!</div>
                                    </div>
                                </div>
                                <div class="col-md-10">
                                    <div class="product-content">
                                        <h5 class="c_left_topplace_tlt">
                                            <a href="#">
                                                <%= db.Get("ProductCode") %>
                                            </a>
                                        </h5>
                                        <span class="tagline c_left_topplace_tagline"><%= db.Get("ProductName") %></span>
                                        <div class="prodect-cotent-image">
    <% 
        if (obj.GetRequest("no") == "1") { 
    %>
                                                <a href="/images/products/hires/<%= productCode %>.jpg" style="border:none" id="hireslink">
                                                    <img src="/images/products/detail/<%= productCode %>.jpg" title="Click for High-Resolution" style="cursor:pointer" alt="<%= db.Get("ProductName") %>" id="productimage" /><br /></a>
    <% 
        } else if (obj.GetRequest("no") != "") { 
    %>
                                                <a href="/images/products/hires/<%= productCode %>-<%= obj.GetRequest(" no ") %>.jpg" style="border:none" id="hireslink">
                                                    <img src="/images/products/detail/<%= productCode %>-<%= obj.GetRequest("no") %>.jpg" title="Click for High-Resolution" style="cursor:pointer" alt="<%= db.Get("ProductName") %>" id="productimage" /><br /></a>
    <% 
        } else { 
    %>
                                                <a href="/images/products/hires/<%= productCode %>.jpg" style="border:none" id="hireslink">
                                                    <img src="/images/products/detail/<%= productCode %>.jpg" title="Click for High-Resolution" style="cursor:pointer" alt="<%= db.Get("ProductName") %>" id="productimage" /><br /></a>
    <% 
        } 
    %>
                                        </div>
                                        <div class="prodect-cotent-buttom">
                                            <span class="c_left_cnt_rightbottom_spn">With Screw Top Lid</span>
                                            <span class="prodect-cotent-buttom-image"><img src="images/imprint_colors.png" alt=""></span>
                                        </div>
                                    </div>
                                </div>
                            </div>
    <% //' *** IF THEY ARE LOGGED IN
        if (Util.GetSession("user") != "") 
        {	
    %>

                            <!-- pricing - calculator button & other pricing -->
    <% 
            if(productCode != "LAN2") 
            { 
    %>
                            <%-- <div id="pricingCalculatorDiv">
                                <button class="btnPriceCalculator" onclick="calculatePriceDialog('<%= productCode %>', '
                            <%= db.Get("ProductName") %>');"><img src="images/price_calculator_icon.jpg" alt="Pricing calculator icon" />Click here for<br>Pricing Calculator</button>
                            --%>
                            <a target="_new" class="btn btn-xlarge btn-info" style="color:blue; background-color: #eee; border: 1px solid #ccc;" href="Quote.aspx?ProductCode=<%= productCode %>">
                                <img align="left" style="margin-top: -18px;margin-right: 30px;" src="images/price_calculator_icon.jpg" width="100" alt="Pricing calculator icon" /><span>Click here for Pricing Calculator</span>
                            </a>
                            <%-- </div> --%>
    <% 
            } 
    %>

                            <div class="ml10">
                                <ul>
                                    <li style="list-style-type:none;"><a style="font-weight:bold;font-size:1.1em;color:#ED3225;" href="#" onclick="popupCenter('/FreightPopup.aspx?id=<%= FreightCodeId %>', 'Freight Pricing', 800, 500); return false;">+ Click for Freight Pricing</a></li>
    <% 
            if (productCode.StartsWith("LAN")) 
            { 
    %>
                                    <li style="list-style-type:none;"><a style="font-weight:bold;font-size:1.1em;color:#ED3225;" href="#" onclick="popupCenter('/LanyardPopup.aspx?id=<%= pid %>'); return false;">+ Click for Lanyard Attachment pricing</a></li>
    <% 
            }
            if (productCode == "B715") 
            { 
    %>
                                    <li style="list-style-type:none;"><a style="font-weight:bold;font-size:1.1em;color:#ED3225;" href="#" onclick="popupCenter('/B715.aspx?id=<%= pid %>', 'Lid Pricing', 600, 350); return false;">+ Click here for Additional Lid Pricing</a></li>
    <% 
            } 
    %>
                                </ul>
                            </div>
                            <!-- /pricing - calculator button & other pricing -->

                            <!-- full pricing available??? -->
    <% 
            if( isFullColourPrint || isScreenPrint || isPadPrint) 
            { 
    %>
                            <hr style='border:1px dotted #aaa;' />

    <% 
                if( isFullColourPrint ) 
                { 
    %>
                            <div class="printingAvailability">
                                <div>
                                    <img src="images/products/printingAvailability/fullColourPrint.png" alt="Full Colour Print Available" />
                                </div>
                                <span>Full Colour Print</span><br />
                                <span>Available</span>
                                <img class="greenTick" src="images/green_tick.png" alt="Green Tick" />
                            </div>
    <% 
                } 
                if( isScreenPrint ) 
                { 
    %>
                            <div class="printingAvailability">
                                <div>
                                    <img src="images/products/printingAvailability/screenPrint.png" alt="Screen Print Available" />
                                </div>
                                <span>Screen Print</span><br />
                                <span>Available</span>
                                <img class="greenTick" src="images/green_tick.png" alt="Green Tick" />
                            </div>
    <% 
                } 
                if( isPadPrint ) 
                { 
    %>
                            <div class="printingAvailability">
                                <div>
                                    <img src="images/products/printingAvailability/padPrint.png" alt="Pad Print Available" />
                                </div>
                                <span>Pad Print</span><br />
                                <span>Available</span>
                                <img class="greenTick" src="images/green_tick.png" alt="Green Tick" />
                            </div>
    <% 
                } 
                if( isWrapPrint ) 
                { 
    %>
                            <div class="printingAvailability">
                                <div>
                                    <img src="images/products/printingAvailability/wrapPrint.png" alt="Wrap Print Available" />
                                </div>
                                <span>Wrap Print</span><br />
                                <span>Available</span>
                                <img class="greenTick" src="images/green_tick.png" alt="Green Tick" />
                            </div>
    <% 
                } 
    %>

    <% 
            }
        } 
    %>
                            <!-- /full pricing available??? -->
                            <div class="row c_left_middle_maketing_tool">
                                <div class="c_left_middle_maketing_tool_tlt">Marketing Tools</div>
                                <div class="col-xs-12 col-md-6 maketing_tool_group">
                                    <button class="btn btn-tool">Quick Stock Check</button>
                                    <button class="btn btn-tool">Freight Estimator</button>
                                </div>
                                <div class="col-xs-12 col-md-6 maketing_tool_group">
                                    <button class="btn btn-tool">Download Images</button>
                                    <button class="btn btn-tool">Create eFlyer</button>
                                </div>
                                <div class="col-xs-12 col-md-6 maketing_tool_group">
                                    <button class="btn btn-tool">Instant Virtual</button>
                                    <button class="btn btn-tool">Add to eCatalog</button>
                                </div>
                                <div class="col-xs-12 col-md-6 maketing_tool_group">
                                    <button class="btn btn-tool">Create eCatalog</button>
                                    <button class="btn btn-tool">Art Template</button>
                                </div>
                            </div>
                            <div class="c_left_bottom_related_product">
                                <div class="c_left_bottom_related_product_tlt">
                                    <span class="glyphicon glyphicon-triangle-left"></span> Related Products
                                    <span class="glyphicon glyphicon-triangle-right"></span>
                                </div>
                                <div id="product-container">
                                    <!-- Suggested products / trending / you may also like / ... -->

                                    <div id="product-suggestions" style="min-height:100px; width:550px; margin-top:30px;">
    <%
        List<string> listProductSuggestionsPageBottom = new List<string>();

        path = Server.MapPath("images/products/thumbnail/");
        DBHelper dbHelperSuggestedProductsPageBottom = getProductSuggestionsPageBottom(pid);

        Response.Write("<div id='productSuggestionsPageBottom'>");
        Response.Write("<div id='productSuggestionsPageBottomInnerDiv'>");
        int nNumSuggestedProducts = dbHelperSuggestedProductsPageBottom.Count;
        foreach (DataRow row in dbHelperSuggestedProductsPageBottom.Rows)
        {
            //if(File.Exists(path + row["ProductCode"] + "-" + i + ".jpg") )
            //{ 
            var html = "<div class='productSuggestions_productWrapper'><div class='productSuggestions_productImage'>";
            html += "<a href='?pid=" + row["ProductId"] + "&code=" + row["ProductCode"] + "' title='" + row["ProductName"] + "' style='margin-top:5px;'><img src='/images/products/thumbnail/" + row["ProductCode"] + ".jpg' style='height:150px; width:150px' /></a>";
            html += "</div><hr/><div class='productSuggestions_productCaption'>";
            html += "<strong>Product Code:</strong> " + row["ProductCode"] + "<br>";
            html += "<strong>Description:</strong> " + row["ProductName"];
            html += "</div>";
            html += "</div>";

            listProductSuggestionsPageBottom.Add("\"" + html + "\"");

            //Response.Write("<a href='?pid=" + row["ProductId"] + "&code=" + row["ProductCode"] + "&no=" + i + "' title='"+ row["ProductName"] +"'><img src='/images/products/thumbnail/" + row["ProductCode"] + ".jpg' style='height:75px; width:75px' /></a>");
            //}
            //}
        } // foreach

        Response.Write("</div></div>");
    %>
                                    </div>
                                    <!-- End Suggested products / trending / you may also like / ... -->
                                </div>
                                    <!-- End Distributor Content -->
                                <!--<div class="col-md-3">
                                    <img src="images/products/B701.png" alt=""><br/> 750ml Sports Bottle - Fliptop<br/> BPA Free<br/> Product Id : B701<br/><br/> As low as:<br/>
                                    <span>$6.17 (AUD)</span><br/>
                                </div>
                                <div class="col-md-3">
                                    <img src="images/products/B703.png" alt=""><br/> 320ml Sports Bottle<br/> BPA Free<br/> Product Id : B703<br/><br/> As low as:<br/>
                                    <span>$6.17 (AUD)</span><br/>
                                </div>
                                <div class="col-md-3">
                                    <img src="images/products/B704.png" alt=""><br/> 500ml Sports Bottle - Fliptop<br/> BPA Free<br/> Product Id : B704<br/><br/> As low as:<br/>
                                    <span>$6.17 (AUD)</span><br/>
                                </div>
                                <div class="col-md-3">
                                    <img src="images/products/B705.png" alt=""><br/> 700ml Dumbbell Sports <br/> Bottle Fliptop BPA Free<br/> Product Id : B705<br/><br/> As low as:<br/>
                                    <span>$6.17 (AUD)</span><br/>
                                </div>-->
                            </div>
                        </div>
                        <!-- /.col-md-6 -->
                        <div class="col-md-6">
                            <div class="product-holder c_right_div">
                                <div class="c_right_child_div">
                                    <div class="c_right_top_part">
                                        <div class="c_r_t_r_tlt">
                                            <%= db.Get("ProductName") %>
                                        </div>
                                        <div class="c_r_t_r_stlt">
                                            <%= db.Get("ProductCode") %>
                                        </div>
                                        <div class="c_r_t_r_price">PRICE (AUG) <img src="images/flg/aus.GIF"></div>
                                    </div>
                                    <table class="product_table">
                                        <thead>
                                            <tr>
                                                <th>In Stock Services</th>
                                                <th>100</th>
                                                <th>250</th>
                                                <th>500</th>
                                                <th>1,000</th>
                                                <th>2,500</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <tr>
                                                <td>7 day Service (C)</td>
                                                <td>$12.93</td>
                                                <td>$12.75</td>
                                                <td>$12.56</td>
                                                <td>$12.38</td>
                                                <td>$12.19</td>
                                            </tr>
                                            <tr>
                                                <td>Set Up (G)</td>
                                                <td>$65.00</td>
                                                <td>$65.00</td>
                                                <td>$65.00</td>
                                                <td>$65.00</td>
                                                <td>$65.00</td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    <table class="product_table">
                                        <thead>
                                            <tr>
                                                <th>Direct Import 10 Weeks</th>
                                                <th>5,000</th>
                                                <th>10,000</th>
                                                <th>25,000</th>
                                                <th>50,000</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <tr>
                                                <td>Ocean Shipping</td>
                                                <td>$12.93</td>
                                                <td>$12.75</td>
                                                <td>$12.56</td>
                                                <td>$12.38</td>
                                            </tr>
                                            <tr>
                                                <td>Set Up (G)</td>
                                                <td>$65.00</td>
                                                <td>$65.00</td>
                                                <td>$65.00</td>
                                                <td>$65.00</td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                                <div class="c_right_child_div">
                                    <div class="c_right_child_tlt">
                                        Product Colours
                                    </div>
                                    <div class="c_right_child_cnt c_r_c_c_clr">
                                        <div class="c_r_c_arrow_l_div c_r_c_clr_div">
                                            <div><span class="glyphicon glyphicon-chevron-left"></span></div>
                                        </div>
                                        <div class="c_r_c_clr_div">
                                            <div style="background: white;">&nbsp;</div>
                                        </div>
                                        <div class="c_r_c_clr_div">
                                            <div style="background: black;">&nbsp;</div>
                                        </div>
                                        <div class="c_r_c_clr_div">
                                            <div style="background: #eaeadd;">&nbsp;</div>
                                        </div>
                                        <div class="c_r_c_clr_div">
                                            <div style="background: #fcdc00;">&nbsp;</div>
                                        </div>
                                        <div class="c_r_c_clr_div">
                                            <div style="background: #d8262f;">&nbsp;</div>
                                        </div>
                                        <div class="c_r_c_clr_div">
                                            <div style="background: #011689;">&nbsp;</div>
                                        </div>
                                        <div class="c_r_c_arrow_r_div c_r_c_clr_div">
                                            <div><span class="glyphicon glyphicon-chevron-right"></span></div>
                                        </div>
                                    </div>
                                </div>
                                <div class="c_right_child_div">
                                    <div class="c_right_child_tlt">
                                        Product Description
                                    </div>
                                    <div class="c_right_child_cnt c_r_c_c_desc">
                                        <%= db.Get("ProductDescription") %>
                                    </div>
                                </div>
                                <div class="c_right_child_div">
                                    <div class="c_right_child_tlt">
                                        Additional Info
                                    </div>
                                    <div class="c_right_child_cnt c_r_c_c_addinfo">
                                        <table class="product_table">
                                            <tr>
                                                <th>Price Includes:</th>
                                                <td>One color one location imprint</td>
                                            </tr>
                                            <tr>
                                                <th>Setup Charge:</th>
                                                <td>$65(G) per color/location</td>
                                            </tr>
                                            <tr>
                                                <th>Repeat Setup Charge:</th>
                                                <td>$65(G) per color(Waived for 2 years)</td>
                                            </tr>
                                            <tr>
                                                <th>Run Charge:</th>
                                                <td>--</td>
                                            </tr>
                                            <tr>
                                                <th>Item Size:</th>
                                                <td>33/8"w X 53/4"l</td>
                                            </tr>
                                            <tr>
                                                <th>Imprint Method:</th>
                                                <td>Pad</td>
                                            </tr>
                                            <tr>
                                                <th>Imprint Area:</th>
                                                <td>(Standard-Pad):21/2"w X 3"h(1 Color Max)<br/> (Optional-Pad: 13/4"w X 21/2"h(2+ Color Max))
                                                </td>
                                            </tr>
                                            <tr>
                                                <th>Optional Imprint Location:</th>
                                                <td>--</td>
                                            </tr>
                                            <tr>
                                                <th>Additional Process:</th>
                                                <td>Pad (Add $0.38(G) per unit for additional locations)</td>
                                            </tr>
                                            <tr>
                                                <th>Packaging Method:</th>
                                                <td>Customized black presentation box</td>
                                            </tr>
                                            <tr>
                                                <th>Carton Dimension (L X W X H):</th>
                                                <td>19" l X 12" w X 13" h</td>

                                            </tr>
                                            <tr>
                                                <th>Weight per Carton(lbs):</th>
                                                <td>18</td>
                                            </tr>
                                            <tr>
                                                <th>Quality per Carton:</th>
                                                <td>30</td>
                                            </tr>
                                            <tr>
                                                <th>Production Time:</th>
                                                <td>10 Weeks, 7 Day Service</td>
                                            </tr>
                                            <tr>
                                                <th>FOB:</th>
                                                <td>Glending NSW 2761</td>
                                            </tr>
                                        </table>
                                    </div>
                                </div>
                                <div class="clearfix"></div>
                            </div>
                            <!-- /.product-holder -->
                        </div>
                        <!-- /.col-md-6 -->
                    </div>
                    <!-- /.row -->
                </div>
                <!-- /.container -->
            </div>
            <!-- /.content-section -->
        </div>

        <script src="js/vendor/jquery-1.10.1.min.js"></script>
        <script src="js/product.js"></script>

        <!--#include file="footerNEW.aspx"-->

        <!-- added by Rian Chipman July 2016 -->
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

            public DBHelper getProductSuggestionsSideBar(string pid) {
                var sql = @"select ProductId, ProductCode, ProductName from Product where Product.ProductId in (select ProductId from ProductInSubCategory where ProductInSubCategory.InSubCatId in (select InSubCatId from ProductInSubCategory where ProductId = " + pid + "))";
                return new DBHelper(sql);
            }
            public DBHelper getProductSuggestionsPageBottom(string pid) {
                var sql = @"SELECT DISTINCT Product.ProductId, ProductCode, ProductName FROM Product 
                INNER JOIN ProductInSubCategory ON ProductInSubCategory.ProductId = Product.ProductId
                WHERE Product.ProductId != " + pid + @"
                AND ProductInSubCategory.InSubCatId IN(select InSubCatId from ProductInSubCategory where ProductId = " + pid + @")
                ORDER BY ProductName ";
                    //Response.Write(sql.Replace("\n", "<br>"));
                return new DBHelper(sql);
            }
        </script>

        <!-- for product suggestions - added by Clinton 15 July, 2016 -->
        <script type="text/javascript">
            <%
                Response.Write("var nNumSuggestedProducts = " + nNumSuggestedProducts + ";");
                Response.Write("var AllSuggestedProducts = [];");

                foreach( string productSuggestionPageBottom in listProductSuggestionsPageBottom  )
                {
                    Response.Write("AllSuggestedProducts.push(" + productSuggestionPageBottom + ");\n");
                }
        /*
                foreach( string productSuggestionProductCode in listProductSuggestionsProductCodes  )
                {
                    Response.Write("AllSuggestedProducts.push(" + productSuggestionProductCode + ");\n");
                }
        */
            %>
        </script>
        <script type="text/javascript" src="js/productSuggestions.js"></script>
        <!-- / for product suggestions - added by Clinton 15 July, 2016 -->