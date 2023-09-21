<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("EZpollAdmin")("name")
	
	If strCookies = "" Then

		Response.Redirect "admin_login.asp"
  
	End If

	intPollID = 0
	If Trim(Request.Form("pollid")) <> "" Then intPollID = checkint(Trim(Request.Form("pollid")))
	
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <div class="row">
        <div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<header><h2>View a poll</h2></header>
        </div>
	</div>
	<form action="admin_viewpoll.asp?l=y" method="post">
	<div class="row">
		<div class="-3u 4u 12u$(medium)" style="padding-bottom:10px;">
		    <% Call selectPoll(intPollID) %>
		</div>
		<div class="-2u$ 12u$(medium)" style="padding-bottom:10px;">
		    <input class="button fit" type="submit" value="View Poll">
		</div>
	</div>
	</form>
<%
	If intPollID <> 0 Then

		intPollLength = 0
		strPollName = ""
		strPollQuestion = ""
		strPollOpenUntil = ""
		datStartDate = Cdate("01/01/1970")
		datEndDate = Cdate("01/01/1970")
		blnHideResults = 0
		blnPollRevote = 0
		blnActive = 0
		strRevote = ""
		strHideResults = ""
		strActive = ""

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"poll WHERE pollID = "&intPollID
		Call getTextRecordset(strSQL,rsCommon)

		If Not rsCommon.EOF Then
			intPollLength = rsCommon("poll_length")
			strPollName = DBDecode(rsCommon("poll_name"))
			strPollQuestion = DBDecode(rsCommon("poll_question"))
			blnHideResults = rsCommon("hide_results")
			blnPollRevote = rsCommon("poll_revote")
			datStartDate = rsCommon("start_date")
			blnActive = rsCommon("poll_active")
		End If
		rsCommon.close
	
		If Cint(intPollLength) > 0 Then
			datEndDate = dateAdd("d",intPollLength,datStartDate)
			strPollOpenUntil = "(Open Until "&datEndDate&")"
		End If		
		
		strSQL = "SELECT Sum(votes) as total FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID
		Call getTextRecordset(strSQL,rsCommon)	
		If Not rsCommon.EOF Then
			total = rsCommon("total")
		End If
		rsCommon.close

		If blnPollRevote Then
			strRevote = "ON"
		Else
			strRevote = "OFF"
		End If

		If blnHideResults Then
			strHideResults = "ON"
		Else
			strHideResults = "OFF"
		End If

		If blnActive Then
		  strActive = "Active"
		Else
		  strActive = "Not Active"
		End If
%>
  <div class="row">
    <div class="-3u 6u 12u(medium)" style="padding-bottom:10px;">
      <div class="row">
		<div class="12u$">
		    <header style="text-align:center;"><h2><%= strPollName %></h2></header>
		    <h4 style="text-align:center;"><%= strPollQuestion %><br /><span style="font-size:12px;"><%= strPollOpenUntil %></span></h4>  
		</div>
        <div class="4u 12u(medium)">
          This Poll is <span class="first" style="color:#FF0000;"><%= strActive %></span>
        </div>
        <div class="4u 12u(medium)">
          Re-vote is <span style="color:#FF0000;"><%= strRevote %></span>
        </div>
        <div class="4u$ 12u(medium)">
          Hide Results is <span style="color:#FF0000;"><%= strHideResults %></span>      
        </div>
      </div>
    </div>
    <div class="-3u 6u 12u(medium)" style="padding-bottom:10px;">
    </div>
<%
		intCounter = 0

		strSQL = "SELECT * FROM "&msdbprefix&"poll_choices WHERE pollID = "&intPollID&" ORDER BY choiceID asc"
		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			Do While Not rsCommon.EOF
				intCounter = intCounter+1
%>			
    <div class="-3u 6u 12u(medium)" style="margin-bottom:-5px">
      <span class="first" style="font-size:14px"><%= intCounter %>. <%= DBDecode(rsCommon("pAnswer")) %>&nbsp;&nbsp;<span class="first" style="font-size:10px"><%= rsCommon("votes") %> (<%= totalCountC(rsCommon("votes"),total) %>)</span></span>
    </div>
    <div class="-3u 6u 12u(medium)" style="padding-bottom:10px;">
      <img src="/poll/images/Image1.jpg" style="height:10px;width:<%= totalCount(rsCommon("votes"),total) %>;border:0px;" />
    </div>
<%
				rsCommon.MoveNext
				If rsCommon.EOF Then Exit Do
			Loop
		End If
		Call closeRecordset(rsCommon)
	    Call ConnClose(Conn)
%>
    <div class="-3u 6u 12u(medium)">
      <span class="first" style="font-size:10px">There Are <%= total %> Votes!</span>
    </div>
  </div>
<% End If %>
</div>
<!-- #include file="../includes/footer.asp"-->