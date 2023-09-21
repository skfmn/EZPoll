<!-- #include file="../includes/general_includes.asp"-->
<%
on error resume next
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

    Set Conn = Server.CreateObject("ADODB.Connection")
    Call ConnOpen(Conn)

	If Trim(Request.QueryString("action")) = "createpoll" Then

		intPollID = 0
		intPollLength = 0
		strPollName = ""
		strPollQuestion = ""
		blnHideResults = 0
		blnPollRevote = 0
		blnActive = 0

		strPollName = Trim(Request.Form("pollname"))
		strPQuestion = Trim(Request.Form("pquestion"))
		blnHideResults = Trim(Request.Form("poll_hide"))
		blnPollRevote = Trim(Request.Form("poll_revote"))
		blnActive = Trim(Request.Form("poll_active"))

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

		If blnActive = "on" Then
		    blnActive = 1
		Else
		    blnActive = 0
		End If

		If Trim(Request.Form("poll_length")) <> "" then
		    intPollLength = checkInt(Trim(Request.Form("poll_length")))
		End If

		strSQL = "INSERT INTO "&msdbprefix&"poll([poll_name],[poll_question],[hide_results],[poll_revote],[poll_length],[start_date],[poll_lock],[poll_active]) Values('"&DBEncode(strPollName)&"','"&DBEncode(strPQuestion)&"',"&blnHideResults&","&blnPollRevote&","&intPollLength&",'"&now&"',0,"&blnActive&")"
		Call getExecuteQuery(strSQL)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		Call getTableRecordset(msdbprefix&"poll",rsCommon)
		If Not rsCommon.EOF Then
		    rsCommon.MoveLast
		    intPollID = Cint(rsCommon("pollID"))
		End If
		Call closeRecordset(rsCommon)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"poll WHERE pollID = "&intPollID

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			If rsCommon("poll_active") <> blnActive Then

				strSQL = "UPDATE "&msdbprefix&"poll SET poll_active = 0 WHERE poll_active = 1"
				Call getExecuteQuery(strSQL)

				strSQL = "UPDATE "&msdbprefix&"poll SET poll_active = 1 WHERE pollID = "&intPollID
				Call getExecuteQuery(strSQL)

			End If
		End If
		Call closeRecordset(rsCommon)

		For Each Item in Request.Form
			If left(Item,7) = "options" Then

			  strPostPollAnswer = Request.Form(Item)
			  intChoiceID = Cint(right(Item,Len(Item)-7))

			  strSQL = "INSERT INTO "&msdbprefix&"poll_choices([pollID],[choiceID],[pAnswer],[votes]) Values("&intPollID&","&intChoiceID&",'"&DBEncode(strPostPollAnswer)&"',0)"
			  Call getExecuteQuery(strSQL)

			End If
	    Next

	    Response.Cookies("msg") = "pcr"
		Response.Redirect "admin_createpoll.asp"

	End If

    Call ConnClose(Conn)
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">

            <header>
                <h2>Create a poll</h2>
            </header>

    <form action="admin_createpoll.asp?action=createpoll" method="post" name="reply">
        <div class="row">
            <div class="6u 12u$(medium)">
                <label for="pollname" style="margin-bottom: -3px;">Poll Name:</label>
                <input type="text" id="pollname" name="pollname" size="50" />
				<br />

                <label for="pquestion" style="margin-bottom: -3px;">Poll Question:</label>
                <input type="text" id="pquestion" name="pquestion" size="50" />
				<br />

                <h4 style="margin-bottom: -5px;">Poll Answers</h4>
                <label for="options-0" style="margin-bottom: -3px;">Answer 1</label>
                <input type="text" name="options1" id="options-0" tabindex="1" size="20" value="" />
				<br />

                <label for="options-1" style="margin-bottom: -3px;">Answer 2</label>
                <input type="text" name="options2" id="options-1" tabindex="2" size="20" value="" />
				<br />

                <span id="pollMoreOptions"></span><a class="button fit nohijack" href="javascript:addPollOption(); void(0);">Add an Answer</a>
            </div>
            <div class="6u$ 12u$(medium)">

                <h4>Show Results</h4>
                <input type="radio" id="poll_hide1" name="poll_hide" value="1" checked="checked" />
                <label for="poll_hide1">Show the poll's results to anyone.</label>
                <input type="radio" id="poll_hide2" name="poll_hide" value="0" />
                <label for="poll_hide2">Only show the results after someone has voted.</label>

                <h4>Allow posters to change vote:</h4>
                <input type="radio" id="poll_revotey" name="poll_revote" value="1" checked="checked" />
                <label for="poll_revotey">Yes</label>
                <input type="radio" id="poll_revoten" name="poll_revote" value="0" />
                <label for="poll_revoten">No</label>

                <h4>Length of poll in days</h4>
                <input type="text" id="poll_length" name="poll_length" style="width: 100px;" />
                <label for="poll_length">leave blank for no close date.</label>

                <input type="checkbox" id="poll_active" name="poll_active" />
                <label for="poll_active">Set This poll as active:</label>

                <input class="button fit" type="submit" name="submit" value="Create Poll" />
            </div>
        </div>

    </form>
</div>
<!-- #include file="../includes/footer.asp"-->