<% 
on error resume next
    Sub trace(strText)
	    Response.Write "Debug: "&strText&"<br />"&vbcrlf
    End Sub

    Sub catch(sText,sText2)

	    If Err.Number <> 0 then
		    Call trace(sText&" - "&err.description)
	    Else
		    Call trace(sText&" - no error")
	    End If
	    If sText2 <> "" Then
		    Call trace(sText&" - "&sText2)
	    End If

	    on error goto 0
    End Sub

    Const ForReading = 1
    Const TristateUseDefault = -2

	Function DBEncode(DBvalue)
		Dim fieldvalue
		fieldvalue = DBvalue

		If fieldvalue <> "" AND Not IsNull(fieldvalue) Then
		
			Set encodeRegExp = New RegExp
			encodeRegExp.Pattern = "((delete)*(select)*(update)*(into)*(drop)*(insert)*(declare)*(xp_)*(union)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
				fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue=replace(fieldvalue,"'","''")

		End If

		DBEncode = fieldvalue

	End Function    
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Install</title>
<link type="text/css" rel="stylesheet" href="../assets/css/main.css" />
</head>
<body>
  <div id="main" class="container" align="center" style="margin-top:-75px;">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <header><h2>EZPoll Installation</h2></header>
      </div>
    </div>
  </div>
<% If Trim(Request.QueryString("step")) = "one" Then %>
  <div id="main" class="container" align="center" style="margin-top:-100px;">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?setsql=y" method="post">
        
        <header>
          <h2>MSSQL Database</h2>
        </header>
        <div class="row">
          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="svrname" style="text-align:left;">Server Host Name or IP Address
              <input type="text" name="svrname" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbname" style="text-align:left;">Database Name
              <input type="text" name="dbname" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbid" style="text-align:left;">Database Login
              <input type="text" name="dbid" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbpwd" style="text-align:left;">Database Password
              <input type="password" name="dbpwd" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbprefix" style="text-align:left;">Table Prefix
              <input type="text" name="dbprefix" value="ezpoll_" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>     
      </div>
    </div>
  </div>
<% 
  ElseIf Request.QueryString("setsql") = "y" Then

%>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
<%

    msdbserver = Trim(Request.Form("svrname"))
    msdb = Trim(Request.Form("dbname"))
    msdbid = Trim(Request.Form("dbid"))
    msdbpwd = Trim(Request.Form("dbpwd"))
    msdbprefix = Trim(Request.Form("dbprefix"))

    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

    Response.Write "Creating Database Tables<br /><br />"

    Response.Write "Creating admin table...<br />"
    Response.Flush
	
    Conn.Execute "CREATE TABLE "&msdbprefix&"admin " & _
    "([adminID] [numeric](10, 0) IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"admin] PRIMARY KEY, " & _
    "[name] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[pwd] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[salt] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
	 
    Response.Write "Populating admin table...<br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"admin ([name],[pwd],[salt]) VALUES ('admin','EB36FB0C1F1A92A838AA1ECAAD4AB6E3B5257103','833D1')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating settings table...<br />"
    Response.Flush
	
    Conn.Execute "CREATE TABLE "&msdbprefix&"settings "& _
    "([settingID] [numeric](10, 0) IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"settings] PRIMARY KEY, " & _
    "[site_title] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[domain_name] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating Messages table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"messages " & _ 
    "([messageID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"messages] PRIMARY KEY, " & _
    "[msg] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[message] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
  
    Response.Write "Populating Messages table...<br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('pcr','New Poll has been added!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('dlt','The Poll has been Deleted!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('pes','The poll was edited successfully!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('pos','The set active operation was successfully!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('pas','Your vote counted!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('adad','Admin added!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('das','Admin eteledd!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ant','Admin name teken!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nadmin','Not an Admin!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('cpwds','Password changed successfully!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('pwd','Admin Info has been changed!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('mus','Messages saved!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('siu','Site info saved!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('error','An unknown error has occurred<br />Please contact support!')"
		   
    Response.Write "Creating Poll table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"poll " & _ 
    "([pollID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"poll] PRIMARY KEY, " & _
    "[poll_name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[poll_question] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[hide_results] bit NULL, " & _
    "[poll_revote] bit NULL, " & _
    "[poll_length] [numeric] (10, 0) NULL, " & _
    "[start_date]  [smalldatetime] NULL, " & _
    "[poll_lock] bit NULL, " & _
    "[poll_active] bit NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating Poll table...<br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"poll([poll_name],[poll_question],[hide_results],[poll_revote],[poll_length],[start_date],[poll_lock],[poll_active]) VALUES('test Poll','This is a test question',0,1,10,'"&Now&"',0,1)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
		
    Response.Write "Creating Poll Choices table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"poll_choices " & _ 
    "([pollID] [numeric] (10, 0)  NULL, " & _
    "[choiceID] [numeric] (10, 0)  NULL, " & _
    "[pAnswer] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[votes] [numeric] (10, 0) NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
  
    Response.Write "Populating Poll Choices table...<br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"poll_choices([pollID],[choiceID],[pAnswer],[votes]) VALUES(1,1,'Test 1',38)"
    Conn.Execute "INSERT INTO "&msdbprefix&"poll_choices([pollID],[choiceID],[pAnswer],[votes]) VALUES(1,2,'Test 2',142)"
    Conn.Execute "INSERT INTO "&msdbprefix&"poll_choices([pollID],[choiceID],[pAnswer],[votes]) VALUES(1,3,'Test 3',77)"
    Conn.Execute "INSERT INTO "&msdbprefix&"poll_choices([pollID],[choiceID],[pAnswer],[votes]) VALUES(1,4,'Test 4',150)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
		
    Response.Write "Creating database tables...Complete!<br />"
    Response.Flush
						
    Response.Write "<br /><br />"
%>
      </div>
    </div>
  </div>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?step=two" method="post">
        <input type="hidden" name="msdbserver" value="<%= msdbserver %>">
        <input type="hidden" name="msdb" value="<%= msdb %>">
        <input type="hidden" name="msdbid" value="<%= msdbid %>">
        <input type="hidden" name="msdbpwd" value="<%= msdbpwd %>">
        <input type="hidden" name="msdbprefix" value="<%= msdbprefix %>">
        <header>
          <h3><span class="first">You have successfully installed the MSSQL Database<br />Please click the button below to continue</span></h3>
        </header>
        <div class="row">
          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>     
      </div>
    </div>
  </div>
<%  
		Conn.Close: Set Conn = Nothing

  ElseIf Request.QueryString("step") = "two" Then  
%>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?step=three" method="post">
        <input type="hidden" name="msdbserver" value="<%= Trim(Request.Form("msdbserver")) %>">
        <input type="hidden" name="msdb" value="<%= Trim(Request.Form("msdb")) %>">
        <input type="hidden" name="msdbid" value="<%= Trim(Request.Form("msdbid")) %>">
        <input type="hidden" name="msdbpwd" value="<%= Trim(Request.Form("msdbpwd")) %>">
        <input type="hidden" name="msdbprefix" value="<%= Trim(Request.Form("msdbprefix")) %>">
        <input type="hidden" name="PhyPath" value="<%= strPhysPath %>" />
        <header>
          <h2>Path Settings</h2>
        </header>
        <div class="row">
          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbid" style="text-align:left;">Base Directory
              <input type="text" name="bdir" value="<%= Request.ServerVariables("APPL_PHYSICAL_PATH") %>" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dir" style="text-align:left;">EZPoll Directory
              <input type="text" name="dir" value="/poll/" size="40" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>
          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>
      </div>
    </div>
  </div>
<%
    ElseIf Request.QueryString("step") = "three" Then 

        strPageFileName = Server.MapPath("../includes/config.asp")

        Set objPageFileFSO = CreateObject("Scripting.FileSystemObject")

        If objPageFileFSO.FileExists(strPageFileName) Then
        Set objPageFileTs = objPageFileFSO.OpenTextFile(strPageFileName, 2)
        Else
        Set objPageFileTs = objPageFileFSO.CreateTextFile(strPageFileName)
        End If

        strPageEntry = Chr(60) & Chr(37) & vbcrlf & _
        "baseDir=""" & Trim(Request.Form("bdir")) & """" & vbcrlf & _
        "strDir=""" & Trim(Request.Form("dir")) & """" & vbcrlf & _
        "msdbprefix=""" & Trim(Request.Form("msdbprefix")) & """" & vbcrlf & _
        "msdbserver=""" & Trim(Request.Form("msdbserver")) & """" & vbcrlf & _
        "msdb=""" & Trim(Request.Form("msdb")) & """" & vbcrlf & _
        "msdbid=""" & Trim(Request.Form("msdbid") )& """" & vbcrlf & _
        "msdbpwd=""" & Trim(Request.Form("msdbpwd")) & """" & vbcrlf & _
        Chr(37) & Chr(62)
				 
        objPageFileTs.WriteLine strPageEntry
  
        objPageFileTs.Close

        Response.Redirect "install.asp?step=four"

   ElseIf Request.QueryString("step") = "four" Then 
%>
  <div id="main" class="container" style="margin-top:-100px;">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <form action="install.asp?step=five" method="post">
        <header>
          <h2>Other stuff</h2>
        </header>
        <div class="row">

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="sitetitle" style="text-align:left;">Site title
              <input type="text" name="sitetitle" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="domainname" style="text-align:left;">Domain name
              <input type="text" name="domainname" value="<%= Request.ServerVariables("SERVER_NAME") %>" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>      
      </div>
    </div>
  </div>
<%
    ElseIf Request("step") = "five" Then
 %><!-- #include file="../includes/config.asp"--><%
        Set Conn = Server.CreateObject("ADODB.Connection")
        Conn.Open "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

        Conn.Execute "INSERT INTO "&msdbprefix&"settings ([site_title],[domain_name]) VALUES ('"&Trim(Request.Form("sitetitle"))&"','"&Trim(Request.Form("domainname"))&"')"

        Conn.Close: Set Conn = Nothing

        Response.Redirect "install.asp?step=done"

    ElseIf Request("step") = "done" Then
%>
  <div id="main" class="container">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <span class="first">
          Success!
          <br>
          You have successfully configured EZPoll!
          <br>
          The next step is to change your password.
          <br>
          Click on the link below and login to admin.
          <br>
          Click on "Password" in the left options menu and change your password.
          <br><br>
          <a class="first" href="../admin/admin_login.asp">Login</a>
        </span>
      </div>
    </div>
  </div>
<% Else %>
  <div id="main" class="container" style="margin-top:-75px;">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <span class="first">
	      You are about to install EZPoll.
	      <br>
	      Please follow the instructions carefully!
	      <br><br>
	      <input class="button" type="button" onClick="parent.location='install.asp?step=one'" value="Continue">
	      <br><br>
	      </span>      
      </div>
    </div>
  </div>
<% End If %>
<br />
</body>
</html>