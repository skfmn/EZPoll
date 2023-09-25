<!-- #include file="includes/general_includes.asp"-->
<%
on error resume next
	intPollID = 0
	intChoiceID = 0
	total = 0
	
	If Trim(Request.Form("castvote")) = "Yes" Then

		intChoiceID = Trim(Request.Form("choiceid")) 
		intPollID = Trim(Request.Form("pollid"))

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"poll_choices WHERE choiceID = "&intChoiceID&" AND pollID = "&intPollID

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			total = rsCommon("votes")
		End If
		Call closeRecordset(rsCommon)
	
		strSQL = "UPDATE "&msdbprefix&"poll_choices SET votes = "&Cint(total)+1&" WHERE choiceID = "&intChoiceID&" AND pollID = "&intPollID
		Call getExecuteQuery(strSQL)

	    Call ConnClose(Conn)

		Response.Cookies("EZPoll")("pollid"&intPollID) = intChoiceID

	End If
	
	If Trim(Request.QueryString("change")) <> "" Then
	
		blnPollRevote = False
		intPollID = checkint(Trim(Request.QueryString("change")))
		intChoiceID = Request.Cookies("EZPoll")("pollid"&intPollID)

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)
		
		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"poll WHERE pollID = "&intPollID
		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
		    blnPollRevote = rsCommon("poll_revote")
		End If
		Call closeRecordset(rsCommon)
		
		If blnPollRevote Then
		
			strSQL = "UPDATE "&msdbprefix&"poll_choices SET votes = (votes-1) WHERE pollID = "&intPollID&" AND choiceID = "&intChoiceID 
			Call getExecuteQuery(strSQL)
		 
			Response.Cookies("EZPoll")("pollid"&intPollID) = ""
		 	
			Response.Redirect "?show=vote"
			
		End If
	    Call ConnClose(Conn)
		
	End If

	blnVoted = False
	intPollID = 0
	intChoiceID = 0
	datStartDate = ""

	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)
	
	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM "&msdbprefix&"poll WHERE poll_active = 1"
	Call getTextRecordset(strSQL,rsCommon)
	If Not rsCommon.EOF Then
	    intPollID = rsCommon("pollID")
	    datStartDate = rsCommon("start_date")
	    intPollLength = rsCommon("poll_length")
	End If
	Call closeRecordset(rsCommon)
	Call ConnClose(Conn)

	datEndDate = DateAdd("d",intPollLength,datStartDate)

	If Request.Cookies("EZPoll")("pollid"&intPollID) <> "" Then
	    blnVoted = True
	End If

	If Trim(Request.QueryString("show")) = "results" Then

	    Call showResults 

	Else

		If blnVoted Then

		    Call showResults

		Else

		    Call showPoll

		End If

	End If
	 
%>