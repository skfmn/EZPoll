<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("EZpollAdmin")("name")

	If strCookies = "" Then
		Response.Redirect "admin_login.asp"
	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

	intPollID = 0
	If Trim(Request.Form("pollid")) <> "" Then
	    intPollID = checkint(Trim(Request.Form("pollid")))
	End If

	intPID = 0
    intPID = Trim(Request.Cookies("pid"))

	If intPID <> "" Then
		intPollID = intPID
        Response.Cookies("pid") = ""
	End If

	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	If Trim(Request.Form("edit")) = "yes" Then

		intPollLength = 0
		strPollName = ""
		strPQuestion = ""
		blnHideResults = 0
		blnPollRevote = 0
		blnActive = 0
		datStartDate = ""

		strPollName = Trim(Request.Form("pollname"))
		strPQuestion = Trim(Request.Form("pquestion"))
		blnHideResults = Trim(Request.Form("poll_hide"))
		blnPollRevote = Trim(Request.Form("poll_revote"))

		If Trim(Request.Form("poll_active")) = "on" Then
		    blnActive = 1
		Else
		    blnActive = 0
		End If

		If blnHideResults Then
		    blnHideResults = 1
		Else
		    blnHideResults = 0
		End If

		If blnPollRevote Then
		    blnPollRevote = 1
		Else
		    blnPollRevote = 0
		End If

		If Trim(Request.Form("poll_length")) <> "" then
		    intPollLength = checkInt(Trim(Request.Form("poll_length")))
		End If

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"poll WHERE pollID = "&intPollID

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			If Cint(rsCommon("poll_length")) = Cint(intPollLength) Then
			    datStartDate = rsCommon("start_date")
			Else
			    datStartDate = Now
			End If

			If rsCommon("poll_active") <> blnActive Then
				If rsCommon("poll_active") = 1 Then

					strSQL = "UPDATE "&msdbprefix&"poll SET poll_active = 0 WHERE pollID = "&intPollID
					Call getExecuteQuery(strSQL)

				Else

					strSQL = "UPDATE "&msdbprefix&"poll SET poll_active = 0 WHERE poll_active = 1"
					Call getExecuteQuery(strSQL)

					strSQL = "UPDATE "&msdbprefix&"poll SET poll_active = 1 WHERE pollID = "&intPollID
					Call getExecuteQuery(strSQL)

				End If
			End If
		End If
		Call closeRecordset(rsCommon)

		strSQL = "UPDATE "&msdbprefix&"poll SET poll_name = '"&DBEncode(strPollName)&"', poll_question = '"&DBEncode(strPQuestion)&"', hide_results = "&blnHideResults&", poll_revote = "&blnPollRevote&", poll_length = "&intPollLength&", start_date = '"&datStartDate&"' WHERE pollID = "&intPollID
		Call getExecuteQuery(strSQL)

		For Each Item in Request.Form
			If left(Item,7) = "options" Then

				strPostPollAnswer = Trim(Request.Form(Item))
				intChoiceID = Cint(right(Item,Len(Item)-7))

				Set rsCommon = Server.CreateObject("ADODB.Recordset")
				strSQL = "SELECT * FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID&" AND choiceID = "&intChoiceID

				Call getTextRecordset(strSQL,rsCommon)
				If Not rsCommon.EOF AND strPostPollAnswer <> "" Then

					strSQL = "UPDATE "&msdbprefix&"poll_choices SET pAnswer = '"&DBEncode(strPostPollAnswer)&"' WHERE pollID = "&intPollID&" AND choiceID = "&intChoiceID
					Call getExecuteQuery(strSQL)

				Else

					If strPostPollAnswer = "" Then
						strSQL = "DELETE FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID&" AND choiceID = "&intChoiceID
					Else
					    strSQL = "INSERT INTO "&msdbprefix&"poll_choices([pollID],[choiceID],[pAnswer],[votes]) Values("&intPollID&","&intChoiceID&",'"&DBEncode(strPostPollAnswer)&"',0)"
					End If
					Call getExecuteQuery(strSQL)

				End If
				Call closeRecordset(rsCommon)

			End If
	    Next

	    Response.Cookies("msg") = "pes"
	    Response.Cookies("pid") = intPollID
	    Response.Redirect "admin_editpoll.asp"

	ElseIf Trim(Request.Form("submit")) = "Delete Poll" Then

		strSQL = "DELETE FROM "&msdbprefix&"poll WHERE pollID = "&intPollID
		Call getExecuteQuery(strSQL)

		strSQL = "DELETE FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID
		Call getExecuteQuery(strSQL)


	    Response.Cookies("msg") = "dlt"
		Response.Redirect "admin_editpoll.asp"

	End If
	Call ConnClose(Conn)
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <div class="row">
        <div class="-4u 4u$ 12u(medium)" style="padding-bottom: 10px;">
            <header>
                <h2>Edit a poll</h2>
            </header>
        </div>
    </div>
    <form action="admin_editpoll.asp" method="post">
        <input type="hidden" name="l" value="y" />
        <div class="row">
            <div class="-4u 4u$ 12u$(medium)" style="padding-bottom: 10px;">
                <% Call selectPoll(intPollID) %>
            </div>
            <div class="-4u 4u$ 12u$(medium)" style="padding-bottom: 10px;">
                <input class="button fit" type="submit" value="Edit Poll">
            </div>
        </div>
    </form>
    <%
	If intPollID <> 0 Then

		strChecked = ""

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT "&msdbprefix&"poll.*, "&msdbprefix&"poll_choices.* FROM "&msdbprefix&"poll INNER JOIN "&msdbprefix&"poll_choices ON "&msdbprefix&"poll.pollID  = "&msdbprefix&"poll_choices.pollID WHERE "&msdbprefix&"poll.pollID = "&intPollID&" ORDER BY choiceID asc"

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			strPollName = DBDecode(rsCommon("poll_name"))
			strPollQuestion = DBDecode(rsCommon("poll_question"))
			datStartDate = rsCommon("start_date")
			blnHideResults = rsCommon("hide_results")
			blnPollRevote = rsCommon("poll_revote")
			blnActive = rsCommon("poll_active")
			intPollLength = rsCommon("poll_length")
    %>
    <form action="admin_editpoll.asp" method="post" name="reply">
        <input type="hidden" name="pollid" value="<%= intPollID %>">
        <input type="hidden" name="edit" value="yes" />
        <div class="row">
            <div class="6u 12u$(medium)">

                <label for="pollname" >Poll Name</label>
                <input type="text" id="pollname" name="pollname" size="50" value="<%= strPollName %>" />
				<br />

                <label for="pquestion">Poll Question:</label>
                <input type="text" id="pquestion" name="pquestion" size="50" value="<%= strPollQuestion %>" />
				<br />

                <h4>Poll Answers&nbsp;<span style="color: #FF0000; font-size: 12px">To delete an answer, leave it blank.</span></h4>
<%
			intCounter = 0
			Do While Not rsCommon.EOF
				intCounter = intCounter+1
%>
                        <label for="options-<%= intCounter %>" style="margin-bottom:-3px;" >Answer <%= intCounter %></label>
                        <input type="text" name="options<%= rsCommon("choiceID") %>" id="options-<%= intCounter %>" value="<%= rsCommon("pAnswer") %>" tabindex="<%= intCounter %>" size="20" style="margin-bottom:10px;" />
<%
			    rsCommon.MoveNext
				If rsCommon.EOF Then Exit Do
			Loop
		End If
		Call closeRecordset(rsCommon)
	    Call ConnClose(Conn)
%>
				<span id="pollMoreOptions"></span><a style="margin-top:10px;" class="button fit nohijack" href="javascript:addPollOption(); void(0);">Add an answer</a>

            </div>
            <div class="6u$ 12u$(medium)">

                <h4>Show Results</h4>
                <% If blnHideResults Then %>
                <input type="radio" id="poll_hide1" name="poll_hide" value="0" />
                <label for="poll_hide1">Show the poll's results to anyone.</label>
                <input type="radio" id="poll_hide2" name="poll_hide" value="1" checked="checked" />
                <label for="poll_hide2">Only show the results after someone has voted.</label>
                <% Else %>
                <input type="radio" id="poll_hide1" name="poll_hide" value="0" checked="checked" />
                <label for="poll_hide1">Show the poll's results to anyone.</label>
                <input type="radio" id="poll_hide2" name="poll_hide" value="1" />
                <label for="poll_hide2">Only show the results after someone has voted.</label>
                <% End If %>

                <h4>Allow posters to change vote:</h4>
                <% If blnPollRevote Then %>
                <input type="radio" id="poll_revote1" name="poll_revote" value="1" checked="checked" />
                <label for="poll_revote1">Yes</label>
                <input type="radio" id="poll_revote2" name="poll_revote" value="0" />
                <label for="poll_revote2">No</label>
                <% Else %>
                <input type="radio" id="poll_revote1" name="poll_revote" value="1" />
                <label for="poll_revote1">Yes</label>
                <input type="radio" id="poll_revote2" name="poll_revote" value="0" checked="checked" />
                <label for="poll_revote2">No</label>
                <% End If %>

                <h4>Length of poll in days</h4>
                <input type="text" id="poll_length" name="poll_length" style="width: 75px;" value="<%= intPollLength %>" />
               <label for="poll_length">Entering a new number will reset the poll start date<br />Leave blank for no close date.</label>

                <% If blnActive Then strChecked = " checked=""checked""" %>
                <input type="checkbox" id="poll_active" name="poll_active" <%= strChecked %> />
                <label for="poll_active">Set This poll as active:</label>

                <input class="button fit" type="submit" name="submit" value="Edit Poll" />

                <a class="button fit" onclick="return confirmSubmit('Are you SURE you want to delete this Poll?','admin_editpoll.asp?delete=yes&pid=<%= intPollID %>')" style="cursor:pointer;">Delete</a>

            </div>
        </div>

    </form>
    <% End If %>
</div>
<!-- #include file="../includes/footer.asp"-->