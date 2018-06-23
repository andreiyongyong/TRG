<%@ Application Language="C#" %>
<%@ Import Namespace="System.Timers" %>
<%@ Import Namespace="System.Net.Mail" %>
<%@ Import Namespace="System.Data" %>

<script language="C#" runat="server"> 
	protected void Application_Start(Object sender, EventArgs e)
	{
		Application["smtpServer"]	= "dedrelay.secureserver.net";
		Application["errorEmail"]	= "rianchipman@gmail.com";
		Application["DSN"]			= @"Data Source=(local);Integrated Security=True;Initial Catalog=TheCorporateCollection;Min Pool Size=1;Connection Reset=false";

//		SendMail("rianchipman@gmail.com", "rianchipman@gmail.com>", "Starting Up", "The Application Pool just re-started");

//		System.Timers.Timer myTimer = new System.Timers.Timer(120 * 1000); // 2 minute in ms
//		myTimer.Elapsed += new ElapsedEventHandler(TimerTest);
//		myTimer.Enabled = true;
	}

	public static void SendMail(string toList, string from, string subject, string msg)
	{
		string mailServer = HttpContext.Current.Application["smtpServer"].ToString();
		System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient(mailServer);
		System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(from, toList, subject, msg);
		message.IsBodyHtml = true;
		try
		{
			client.Send(message);
		}
		catch (Exception e) 
		{
			HttpContext.Current.Response.Write("ERROR: " + e.Message);
		}
	}
	
	private static void TimerTest(object source, ElapsedEventArgs e)
	{
		// Can't access Application variables yet
		//string mail = HttpContext.Current.Application["errorEmail"].ToString();
		//SendMail(mail, mail, "Test Timer", "Timers are like Windows Services than run all the time!");
	}

	
	void Application_Error(Object sender, EventArgs e)
	{
	return;
		if (Response.Buffer == true) Response.Clear();
			string url = "http://" + Request.ServerVariables["SERVER_NAME"] + Request.ServerVariables["SCRIPT_NAME"];
		if (Request.ServerVariables["QUERY_STRING"] != "")
			url = url + "?" + Request.ServerVariables["QUERY_STRING"];
		else if (Request.Form.ToString() != "")
			url += "?" + Request.Form.ToString();

		if (url.Contains("trace.axd") || url.Contains(".rem") || url.Contains(".soap")) return;
		
		Exception oErr = Server.GetLastError(); 
		Exception myOwnError = oErr;
		string msg = "";
		string myError = oErr.ToString();

		if (oErr.GetType().ToString() == "System.Web.HttpCompileException") return;
		
		myError = oErr.GetBaseException().ToString();
		if (myError.Contains("does not exist")) return;

		string desc = Between(myError, ": ", "at "); // everything between line and a new line
		myError = Between(myError, " in ", "");
		myError = myError.Replace(":line ", "<br>Line ");
		
        msg += "<table border=1 cellpadding=0 cellspacing=0 style='border-collapse:collapse; border: 1px navy solid'>";
        msg += "<tr><td>Remote IP</td><td>" + Request["REMOTE_ADDR"].ToString() + "</td></tr>";
        msg += "<tr><td nowrap>Server IP</td><td>" + Request["LOCAL_ADDR"].ToString() + "</td></tr>";
		msg += "<tr><td>URL:</td><td>" + url + "</td></tr>";
		msg += "<tr><td>Referer:</td><td>" + Request.ServerVariables["HTTP_REFERER"] + "</td></tr>";
        msg += "<tr><td>Browser:</td><td>" + Request.ServerVariables["HTTP_USER_AGENT"] + "</td></tr>";
		msg += "<tr><td>Type:</td><td>" + oErr.GetBaseException().GetType() + " Top: " + oErr.GetType() + "</td></tr>";
		msg += "<tr><td>Error:</td><td style='border: 2px Solid Red; padding: 2px; background-color: #F08080;'>" + oErr.GetBaseException().Message + "</td></tr>";
		string[] splitup = oErr.GetBaseException().StackTrace.Split('\n');
        string stack = "";
            
        for(int i = 0; i < splitup.Length; i++)
        {
            stack+="<br/>";
            string rp = "";
            for(int k = 0; k < 2*i; k++)
            {
                rp += "&nbsp;";
            } 
            stack += rp;
            if(splitup[i].Contains(" in "))
                {
                splitup[i] = splitup[i].Replace(" in ","\n<font color=maroon>in </font>");
                stack += splitup[i].Split('\n')[0];
                stack += "<br/>&nbsp;" + rp + "<b>";
                stack += splitup[i].Split('\n')[1].Replace("line","<font color=red>line</font>");
                stack += "</b>";
                }
                else
                stack += splitup[i]; 
        }
        msg += " <tr><td>Stack trace:</td><td>" + stack + "</td></tr>";
        msg += "</table>";
		//Response.Write(msg);

		SendMail("rianchipman@gmail.com", "error@godaddy.com", "ASP.NET Error", msg );

		Response.End();
	}
	
	public string LeftOf(string str, string chr) 
	{
		int leftOfPos = str.IndexOf(chr);
		if (leftOfPos <= 0) return str;
		return str.Substring(0, leftOfPos);
	}
	public string RightOf(string str, string searchfor) 
	{
		int pos = str.IndexOf(searchfor);
		if (pos >= 0) return str.Substring(pos + searchfor.Length);
		return "";
	}
	public string Between(string str, string searchStart, string searchEnd) 
	{
		return LeftOf(RightOf(str, searchStart), searchEnd);
	}
</script>
