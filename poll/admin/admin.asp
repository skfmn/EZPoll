<!-- #include file="../includes/general_includes.asp"-->
<%
    strCookies = Request.Cookies("EZPollAdmin")("name")

    If strCookies = "" Then
        Response.Redirect "admin_login.asp"
    End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

    If Request.Form("chpwd") = "y" Then

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM admin"
        Call getTextRecordset(strSQL,rsCommon)
        If Not rsCommon.EOF Then
            rsCommon("admin_name") = DBEncode(Trim(Request.Form("aname")))
            rsCommon("admin_pwd") = md5(Trim(Request.Form("pwd")))
            rsCommon.Update

            Response.Cookies("EZPoll")("un") = rsCommon("admin_name")
            Response.Cookies("EZPoll").Expires = Date + 30
        End If
        Call closeRecordset(rsCommon)
        Call ConnClose(Conn)

    End If

    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(Server.MapPath("/poll/install")) Then
        fso.DeleteFolder(Server.MapPath("/poll/install"))
    End If

    Set fso = Nothing
%>
<!-- #include file="../includes/header.asp"-->
<% If msg <> "" Then Call displayFancyMsg(getMessage(msg)) %>
<div id="main" class="container">
    <header>
        <h1 style="text-align: center;font-size:30px">EZPoll</h1>
        <h4 style="text-align: center;">Choose an Option below</h4>
    </header>
    <div class="row">
        <div class="-3u 3u 12u$(medium)">
            <ul class="alt">
                <li><a class="button fit" href="admin_createpoll.asp"><span>Create Poll</span></a></li>
                <li><a class="button fit" href="admin_editpoll.asp"><span>Edit Polls</span></a></li>
            </ul>
        </div>
        <div class="3u$ 12u$(medium)">
            <ul class="alt">
                <li><a class="button fit" href="admin_viewpoll.asp"><span>View Polls</span></a></li>
                <li><a class="button fit" href="admin_manage.asp"><span>Manage Admins</span></a></li>
            </ul>
        </div>
        <div class="-3u 6u 12u$(medium)">
            <%= getResponse("http://www.aspjunction.com/gnews.asp?ref=y&amp;pv="& strVersion&"") %>
        </div>
    </div>
</div>
<!-- #include file="../includes/footer.asp"-->