<%
    strVersion = "4.0"

    ConnStr = "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

	Function getResponse(sURL)
		Dim strTemp
		strTemp = ""

		Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		xmlhttp.SetOption(2) = (xmlhttp.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
		xmlhttp.Open "GET", sURL, false
		xmlhttp.Send(NULL)
		If xmlhttp.readyState = 4 Then
			If xmlhttp.status <> 200 Then

			    strTemp = "<span style=""color:#FF0000"">Error "&xmlhttp.status&" - "&xmlhttp.statusText&"</span><br />"

			Else

			    strTemp = xmlhttp.ResponseText

			End If
		End If
		Set xmlhttp = Nothing

		getResponse = strTemp

	End Function

	Sub selectPoll(iPollID)
		Dim rsTemp

	    Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		Call getTableRecordset(msdbprefix&"poll",rsTemp)
		If Not rsTemp.EOF Then
			Response.Write "<select id=""pollid"" name=""pollid"">"&vbcrlf
	        Response.Write "<option value="""">Select a Poll</option>"&vbcrlf
			Do While Not rsTemp.EOF
	            If Cint(iPollID) = Cint(rsTemp("pollID")) Then
	                Response.Write "<option value="""&rsTemp("pollID")&""" selected>"&rsTemp("poll_name")&"</option>"&vbcrlf
	            Else
				    Response.Write "<option value="""&rsTemp("pollID")&""">"&rsTemp("poll_name")&"</option>"&vbcrlf
	            End If
				rsTemp.MoveNext
				If rsTemp.EOF Then Exit Do
			Loop
			Response.Write "</select>"&vbcrlf
		Else
			Response.Write "<select id=""pollid"" name=""pollid"">"&vbcrlf
			Response.Write "<option value=""0"">No Polls available</option>"&vbcrlf
			Response.Write "</select>"&vbcrlf
		End If
		Call closeRecordset(rsTemp)
	    Call COnnClose(Conn)

	End Sub

	Sub showPoll
	    Dim rsTemp

		intPollID = 0
		intPollLength = 0
		total = 0
		strPollName = ""
		strPollQuestion = ""
		strPollOpenUntil = ""
		datStartDate = Cdate("01/01/1970")
		datEndDate = Cdate("01/01/1970")
		blnHideResults = 0
		blnPollRevote = 0
		blnActive = 0

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"poll WHERE poll_active = 1"
		Call getTextRecordset(strSQL,rsTemp)
		If Not rsTemp.EOF Then
			intPollID = rsTemp("pollID")
			intPollLength = rsTemp("poll_length")
			strPollName = DBDecode(rsTemp("poll_name"))
			strPollQuestion = DBDecode(rsTemp("poll_question"))
			blnHideResults = rsTemp("hide_results")
			blnPollRevote = rsTemp("poll_revote")
			datStartDate = rsTemp("start_date")
			blnActive = rsTemp("poll_active")
		End If
		Call closeRecordset(rsTemp)

		If Cint(intPollLength) > 0 Then
			datEndDate = dateAdd("d",intPollLength,datStartDate)
			strPollOpenUntil = "(Open Until "&datEndDate&")"
		End If

	    Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT Sum(votes) as total FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID
		Call getTextRecordset(strSQL,rsTemp)
		If Not rsTemp.EOF Then
		    total = rsTemp("total")
		End If
		Call closeRecordset(rsTemp)

%>
<div id="main" class="container">
    <header style="text-align: center;">
        <h2><%= strPollName %></h2>
    </header>
    <h4 style="text-align: center;"><%= strPollQuestion %><br />
        <span style="font-size: 12px;"><%= strPollOpenUntil %></span></h4>
    <script type="text/javascript">
        function validateVote() {
            var radios = document.getElementsByName('choiceid')

            for (var i = 0; i < radios.length; i++) {
                if (radios[i].checked) {
                    return document.pollvote.submit();
                }
            };

            alert('Please select an answer!');
            return false;
        }
    </script>

    <form action="" method="post" name="pollvote" id="pollvote" onsubmit="return validateVote();">
        <input type="hidden" name="pollid" value="<%= intPollID %>" />
        <input type="hidden" name="castvote" value="Yes" />
        <div class="row">
            <%
			intCounter = 0

			Set rsTemp = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID&" ORDER BY choiceID asc"
			Call getTextRecordset(strSQL,rsTemp)
			If Not rsTemp.EOF Then
				Do While Not rsTemp.EOF
					intCounter = intCounter+1
            %>
            <div class="-3u 6u$ 12u$(small)" style="padding-bottom: 10px;">
                <input type="radio" name="choiceid" id="choiceid<%= intCounter %>" value="<%= rsTemp("choiceID") %>" required />
                <label for="choiceid<%= intCounter %>"><%= intCounter %>. <%= DBDecode(rsTemp("pAnswer")) %></label>
            </div>
            <%
					rsTemp.MoveNext
					If rsTemp.EOF Then Exit Do
				Loop
			End If
			Call closeRecordset(rsTemp)
			Call ConnClose(Conn)
            %>
            <% If Not blnHideResults Then %>
            <div class="-3u 3u 12u$(small)" style="padding-bottom: 10px;">
                <a class="button" onclick="return validateVote();">Vote</a>
            </div>
            <div class="3u$ 12u$(small)" style="padding-bottom: 10px;">
                <a class="button" href="?show=results">Show results</a>
            </div>
            <% Else %>
            <div class="-3u 6u$ 12u$(small)" style="padding-bottom: 10px;">
                <a class="button" onclick="document.pollvote.submit()">Vote</a>
            </div>
            <% End If %>
        </div>
    </form>
    <div class="row">
        <div class="-3u 3u 12u$(small)" style="padding: 10px 0;">
            <span style="font-size: 16px; font-weight: bold;">Powered By <a href="http://www.aspjunction.com" target="_blank">EZPoll</a></span>
        </div>
        <div class="3u$ 12u$(small)" style="padding: 10px 0;">
            <span style="font-size: 16px; font-weight: bold;">There are <%= total %> votes!</span>
        </div>
    </div>
</div>
<%
	End Sub

	Sub showResults
	    Dim rsTemp

		intPollID = 0
		intPollLength = 0
		strPollName = ""
		strPollQuestion = ""
		strPollOpenUntil = ""
		datStartDate = Cdate("01/01/1970")
		datEndDate = Cdate("01/01/1970")
		blnHideResults = False
		blnPollRevote = False
		blnActive = False

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

        If Trim(Request.QueryString("pollname")) <> "" Then
		  strSQL = "SELECT * FROM "&msdbprefix&"poll WHERE poll_name = '"&DBEncode(Trim(Request.QueryString("pollname")))&"'"
		Else
		  strSQL = "SELECT * FROM "&msdbprefix&"poll WHERE poll_active = 1"
		End If

		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		Call getTextRecordset(strSQL,rsTemp)

		If Not rsTemp.EOF Then
		    intPollID = rsTemp("pollID")
			intPollLength = rsTemp("poll_length")
			strPollName = DBDecode(rsTemp("poll_name"))
			strPollQuestion = DBDecode(rsTemp("poll_question"))
			blnHideResults = rsTemp("hide_results")
			blnPollRevote = rsTemp("poll_revote")
			datStartDate = rsTemp("start_date")
			blnActive = rsTemp("poll_active")
		End If
		rsTemp.close

		If Cint(intPollLength) > 0 Then
			datEndDate = dateAdd("d",intPollLength,datStartDate)
			strPollOpenUntil = "(Open Until "&datEndDate&")"
		End If

		strSQL = "SELECT Sum(votes) as total FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID
		Call getTextRecordset(strSQL,rsTemp)
		If Not rsTemp.EOF Then
			total = rsTemp("total")
		End If
		rsTemp.close

%>
<div id="main" class="container">
    <header style="text-align: center;">
        <h2><%= strPollName %></h2>
    </header>
    <div class="row">
        <div class="-3u 6u$ 12u$(small)" style="padding-bottom: 10px;">
            <h4><%= strPollQuestion %> <span style="font-size: 12px;"><%= strPollOpenUntil %></span></h4>
        </div>
        <%
		intCounter = 0

		strSQL = "SELECT * FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID&" ORDER BY choiceID asc"
		Call getTextRecordset(strSQL,rsTemp)
		If Not rsTemp.EOF Then
			Do While Not rsTemp.EOF
				intCounter = intCounter+1
        %>
        <div class="-3u 6u$ 12u$(small)" style="margin-bottom: -5px">
            <span style="font-size: 14px"><%= intCounter %>. <%= DBDecode(rsTemp("pAnswer")) %>&nbsp;&nbsp;<span style="font-size: 10px"><%= rsTemp("votes") %> (<%= totalCountC(rsTemp("votes"),total) %>)</span></span>
        </div>
        <div class="-3u 6u$ 12u$(small)" style="padding-bottom: 10px;">
            <img src="/poll/images/Image1.jpg" style="height: 10px; width: <%= totalCount(rsTemp("votes"),total) %>; border: 0px;" />
        </div>
        <%
				rsTemp.MoveNext
				If rsTemp.EOF Then Exit Do
			Loop
		End If
		Call closeRecordset(rsTemp)
	    Call ConnClose(Conn)
        %>
        <div class="-3u 6u$ 12u$(small)">
            <% If Request.Cookies("EZPoll")("pollid"&intPollID) = "" Then %>
            <a class="button" href="?show=vote">Vote</a>
            <% ElseIf Request.Cookies("EZPoll")("pollid"&intPollID) <> "" AND blnPollRevote Then %>
            <a class="button" href="?change=<%= intPollID %>">Change Vote</a>
            <% End If %>
        </div>
        <div class="-3u 3u 12u$(small)" style="padding-top: 10px;padding-bottom: 10px;">
            <span style="font-size: 16px; font-weight: bold;">Powered By <a href="http://www.aspjunction.com" target="_blank">EZPoll</a></span>
        </div>
        <div class="3u$ 12u$(small)" style="padding-top: 10px;padding-bottom: 10px;">
            <span style="font-size: 16px; font-weight: bold;">There are <%= total %> votes!</span>
        </div>
    </div>
</div>
<%
	End Sub

	Function totalCount(iPvalue,iTotal)
	  iPvalue = Clng(iPvalue)
		iTotal = Clng(iTotal)
	  If iTotal = 0 or iPvalue = 0 Then
		  totalCount = "1px"
		Else
		  totalCount = Clng(iPvalue/iTotal*100)&"%"
		End If

	End Function

	Function totalCountC(iPvalue,iTotal)

	  iPvalue = Clng(iPvalue)
		iTotal = Clng(iTotal)
	  If iTotal = 0 or iPvalue = 0 Then
		  totalCountC = "0%"
		Else
		  totalCountC = Clng(iPvalue/iTotal*100)&"%"
		End If

	End Function

	Function getMessage(sMsg)
	    Dim rsTemp, strTemp
		strTemp = ""

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

	    Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT message FROM "&msdbprefix&"messages WHERE msg = '"&Trim(sMsg)&"'"

		Call getTextRecordset(strSQL,rsTemp)
		If Not rsTemp.EOF Then
		    strTemp = rsTemp("message")
		Else
		    strTemp = sMsg
		End If
		Call closeRecordset(rsTemp)
	    Call ConnClose(Conn)

		getMessage = strTemp

	End Function

	Function msgTrans(sMsg)

		Dim strTemp: strTemp = ""

		Select  Case sMsg
			case "pcr"
				strTemp = "Poll created:"
			case "dlt"
			    strTemp = "Poll deleted:"
			case "pwd"
			    strTemp = "Admin changed:"
			case "pos"
			    strTemp = "Set active:"
			case "pas"
			    strTemp = "Vote counted:"
			case "pes"
			    strTemp = "Poll edited:"
	        case "mus"
                strTemp = "Messages saved:"
			case "adad"
			    strTemp = "Admin added:"
			case "ant"
			    strTemp = "Admin taken:"
			case "cpwds"
			    strTemp = "Changed password:"
			case "nadmin"
			    strTemp = "Not Admin:"
			case "das"
			    strTemp = "Admin deleted:"
			case "siu"
			    strTemp = "Site info saved:"
			case "error"
			    strTemp = "Generic error:"
			case else
			    strTemp = "TBD:"
		End Select

		msgTrans = strTemp

	End Function

	Function displayFancyMsg(sText)
%>
<div style="display: none">
    <a id="textmsg" href="#displaymsg">Message</a>
    <div id="displaymsg" style="background-color: #ffffff; text-align: left; width: 300px;">
        <div class="left_menu_block">
            <div class="left_menu_top">
                <h2>Message</h2>
            </div>
            <div class="left_menu_center" align="center" style="background-color: #ffffff; padding-left: 0px;"><span style="color: #000000;"><%= sText %></span></div>
            <div class="left_menu_bottom"></div>
        </div>
    </div>
</div>
<%
	End Function

	Function checkInt(iVal)
		Dim intTemp
		intTemp = 0

		If iVal <> "" Then
			If Not IsNumeric(iVal) Then
				Call displayFancyMsg("Input was not a number!")
			Else
				intTemp = iVal
			End If
		Else
		    Call displayFancyMsg(txtInputWasEmpty&"Input was empty!")
		End If

		checkInt = intTemp

	End Function

	Function DBEncode(DBvalue)
		Dim fieldvalue
		fieldvalue = Trim(DBvalue)

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

	Function DBDecode(DBvalue)
		Dim fieldvalue
		fieldvalue = Trim(DBvalue)

		If fieldvalue <> "" AND ( NOT IsNull(fieldvalue) ) Then

			Set encodeRegExp = New RegExp
			encodeRegExp.Pattern = "((eteled)*(tceles)*(etadpu)*(otni)*(pord)*(tresni)*(eralced)*(_px)*(noinu)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
				fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue = replace(fieldvalue,"''","'")

		End If

		DBDecode = fieldvalue

	End Function

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
%>