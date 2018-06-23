<!DOCTYPE html>
<html>

<head>
    <title>Specialty Range</title>
    <meta http-equiv="X-UA-Compatible" content="IE=9" />
    <meta name="description" content="Specialty Range is a range of promotional products, sold by promotional product distributors around Australia." />
    <meta name="robots" content="index, follow, noarchive" />
    <meta name="googlebot" content="noarchive" />

    <meta name="viewport" content="width=device-width" />

    <link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.4/themes/redmond/jquery-ui.css" />
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css" />

    <link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800" />
    <link rel="stylesheet" type="text/css" media="screen" href="css/style_test_test.css" />
    <link rel="stylesheet" href="css/bootstrap.css" />
    <link rel="stylesheet" href="css/normalize.min.css" />
    <link rel="stylesheet" href="css/animate.css" />
    <link rel="stylesheet" href="css/templatemo-misc.css" />
    <link rel="stylesheet" href="css/templatemo-style.css" />
    <link rel="stylesheet" href="css/custom.css" />

    <script type="text/javascript" src="js/script.js"></script>
    <script type="text/javascript" src="js/vendor/modernizr-2.6.2.min.js"></script>

    <script type="text/javascript">
        window.jQuery || document.write('<script src="js/vendor/jquery-1.10.1.min.js"><\/script>')
    </script>
    <script type="text/javascript" src="js/jquery.easing-1.3.js"></script>
    <script type="text/javascript" src="js/bootstrap.js"></script>
    <script type="text/javascript" src="js/plugins.js"></script>
    <script type="text/javascript" src="js/main.js"></script>
</head>

<body>
    <%@ Import Namespace="System.Data" %>
    <%@ Import Namespace="System.IO" %>
    <%
        DBHelper db = null;
        System.Util obj = new System.Util();
        string serverName = Request.ServerVariables["SERVER_NAME"].ToString();
    %>
        <header class="site-header">
            <div class="main-header">
                <div class="container">
                    <div class="row">
                        <div class="col-md-6 col-sm-6 col-xs-12 c_left_for_logo">
                            <div class="logo">
                                <a href="#"><img src="images/spectioaltyrange.png" alt="" /></a>
                                <br /><br />
                            </div>
                            <!-- /.logo -->
                        </div>
                        <!-- /.col-md-6 -->
                        <div class="col-md-6 col-sm-6 col-xs-12 menu-menu c_header_right_div">
                            <div class="social-icons">
                                <ul>
                                    <li>
                                        <a href="#" class="facebook"></a>
                                    </li>
                                    <li>
                                        <a href="#" class="youtube"></a>
                                    </li>
                                    <li>
                                        <a href="#" class="linkedin"></a>
                                    </li>
                                    <li>
                                        <a href="#" class="instagram"></a>
                                    </li>
                                </ul>
                                <div class="clearfix"></div>
                            </div>
                            <div class="search">
                                <form action="#" method="post" id="search_form">
                                    <input type="text" placeholder="Search entire store here..." id="s" />
                                    <input type="submit" id="searchform" value="search" />
                                    <input type="hidden" value="post" name="post_type" />
                                </form>
                            </div>
                            <div class="language">
                                <div id="" class="">
                                    <ul class="nav navbar-nav">
                                        <li class="dropdown">
                                            <a href="#" class="dropdown-toggle" data-toggle="dropdown" role="button" aria-expanded="false"><img id="imgNavSel" src="" alt="..." class="img-thumbnail icon-small"><span id="lanNavSel">ITA</span> <span class="caret"></span></a>
                                            <ul class="dropdown-menu" role="menu">
                                                <li>
                                                    <a id="navIta" href="#" class="language"> <img id="imgNavIta" src="" alt="..." class="img-thumbnail icon-small"><span id="lanNavIta">Italiano</span></a>
                                                </li>
                                                <li>
                                                    <a id="navDeu" href="#" class="language"> <img id="imgNavDeu" src="" alt="..." class="img-thumbnail icon-small"><span id="lanNavDeu">Deutsch</span></a>
                                                </li>
                                                <li>
                                                    <a id="navFra" href="#" class="language"><img id="imgNavFra" src="" alt="..." class="img-thumbnail icon-small"><span id="lanNavFra">Francais</span></a>
                                                </li>
                                                <li>
                                                    <a id="navEng" href="#" class="language"><img id="imgNavEng" src="" alt="..." class="img-thumbnail icon-small"><span id="lanNavEng">English</span></a>
                                                </li>
                                            </ul>
                                        </li>
                                    </ul>
                                </div>
                            </div>
                        </div>
                        <!-- /.col-md-6 -->
                        <div class="btnLogin">
                            <% 	if (Util.GetSession("user") == "") { %>
                            <a id="btnLogin" data-toggle="modal" data-target="#modalLogin">
                                <span class="glyphicon glyphicon-log-in Login"></span>
                                <span>Login</span>
                            </a>                               
                        	<% } else { %>
                            <a id="btnLogout" href="/logout.aspx">
                                <span class="glyphicon glyphicon-log-out Login"></span>
                                <span>Logout</span>
                            </a>                                
	                        <% } %>
                        </div>
                    </div>
                    <!-- /.row -->
                </div>
                <!-- /.container -->
            </div>
            <!-- /.main-header -->
            <div class="main-nav">
                <div class="container">
                    <div class="row">
                        <div class="col-md-12 col-sm-12">
                            <div class="list-menu">
                                <a href="#" class="toggle-menu" state="0">
                                    <!--<i class="fa fa-bars"></i>-->
                                    <span class="glyphicon glyphicon-align-justify"></span>
                                </a>
                                <ul class="c_temp_nav">
                                    <li>&nbsp;</li>
                                </ul>
                                <ul class="menu c_nav_menu">
                                    <li><a href="/">Home</a></li>
                                    <li><a href="/browse.aspx?mid=272">Clearance</a></li>
                                    <li><a href="/browse.aspx?mid=274">New Prodects</a></li>
                                    <li><a href="/browse.aspx?mid=316">Spectials</a></li>
                                    <li><a href="/catalogue.aspx">Pricelist & Catalogue</a></li>
                                    <li><a href="/about.aspx">About</a></li>
                                </ul>
                            </div>
                            <!-- /.list-menu -->
                        </div>
                        <!-- /.col-md-6 -->
                    </div>
                    <!-- /.row -->
                </div>
                <!-- /.container -->
            </div>
            <!-- /.main-nav -->
            <div class="c_footer_top_line"></div>
            <div class="c_footer_topbottom_line"></div>
        </header>

        <!-- login modal dialog -->
        <div id="modalLogin" class="modal fade" role="dialog">
            <div class="modal-dialog modal-sm">
                <!-- Modal content-->
                <div class="modal-content">
                    <form action="/login.aspx" method="post" id="loginform">
                        <div class="modal-header">
                            <button type="button" class="close" data-dismiss="modal">&times;</button>
                            <h4 class="modal-title">Distributor Login</h4>
                        </div>
                        <div class="modal-body">
                            <input id="txtusername" type="text" name="txtusername" value="<%= Util.GetCookie(" ckusername ") %>" placeholder="Username" />
                            <input id="txtpassword" type="password" name="txtpassword" value="" placeholder="Password" />
                        </div>
                        <div class="modal-footer">
                            <button type="submit" class="btn btn-default" style="background-color:red">Login</button>
                        </div>
                    </form>
                </div>
            </div>
        </div>

        <div id="content-outer">
            <div id="content-wrapper" class="container_16">
                <!-- /.site-header -->