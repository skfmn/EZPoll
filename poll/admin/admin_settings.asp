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

    If Request.Form("chmsg") <> "" Then

        For Each i in Request.Form
	        If left(i,8) = "messages" Then

		        strFormValue = Replace(i,left(i,9),"")
		        strFormValue = Replace(strFormValue,right(i,1),"")
                strFormMsg = Request.Form(i)

                strSQL = "UPDATE " & msdbprefix & "messages SET message = '"&DBEncode(strFormMsg)&"' WHERE msg = '"&strFormValue&"'"
                Call getExecuteQuery(strSQL)

	        End If
        Next

        Response.Cookies("msg") = "mus"
        Response.Redirect "admin_settings.asp"

    End If

   If Request.Form("chmstgs") <> "" Then

        strSiteTitle = DBEncode(Request.Form("sitename"))
        strDomainname = DBEncode(Request.Form("domainname"))

        strSQL = "UPDATE " & msdbprefix & "settings SET site_title = '"&strSiteTitle&"', domain_name = '"&strDomainname&"'"
        Call getExecuteQuery(strSQL)

        Response.Cookies("msg") = "siu"
        Response.Redirect "admin_settings.asp"

    End If

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <div class="row">
        <div class="-1u 10u$ 12u$(medium)">
            <header>
                <h2>Manage Settings</h2>
            </header>
        </div>
    </div>
    <div class="row uniform">
        <div class="-1u 10u$ 12u$(medium)">
            <h3>Site Settings</h3>
<% 
    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    Call getTableRecordset(msdbprefix & "settings",rsCommon)
    If Not rsCommon.EOF Then
        strSiteTitle = DBDecode(rsCommon("site_title"))
        strDomainname = DBDecode(rsCommon("domain_name"))
    End If
    Call closeRecordset(rsCommon)
%>
            <form action="admin_settings.asp" method="post">
                <input type="hidden" name="chmstgs" value="y" />
                <div class="row">
                    <div class="4u 12u$(medium)">
                        <label for="sitename">Site Name</label>
                        <input type="text" id="sitename" name="sitename" value="<%= strSiteTitle %>" />
                    </div>
                    <div class="4u 12u$(medium)">
                        <label for="domainname">Domain Name</label>
                        <input type="text" id="domainname" name="domainname" value="<%= strDomainname %>" />
                    </div>
                    <div class="4u$ 12u$(medium)">
                        <label for="submit">&nbsp;</label>
                        <input class="button fit" type="submit" name="submit" value="Save Settings" style="vertical-align:bottom;" />
                    </div>
                </div>
            </form>
        </div>
        <div class="-1u 10u$ 12u$(medium)" style="padding-bottom:10px;">
            <h3>Messages</h3>
<%
    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    Call getTableRecordset(msdbprefix & "messages",rsCommon)
    If Not rsCommon.EOF Then
%>         
            <div class="table-wrapper">
                <form action="admin_settings.asp" method="post">
                    <input type="hidden" name="chmsg" value="y" />
                    <table width="100%">
                        <tbody>
                            <tr>
                                <td>
                                    <table>
<%
        intCounter = 0
        Do While not rsCommon.EOF
            intCounter = intCounter+1
            strTempMsg = DBDecode(rsCommon("msg"))
            strMessage = DBDecode(rsCommon("message"))
%>
                                        <tr>
                                            <td style="width:30%;">
                                                <%= msgTrans(strTempMsg) %>
                                            </td>
                                            <td style="width:70%;">
                                                <input type="text" name="messages[<%= strTempMsg %>]" value="<%= strMessage %>" />
                                            </td>
                                        </tr>
<% 
            If intCounter = 7 Then 
%>

                                    </table>
                                </td>
                                <td>
                                    <table>
<%
            End If

            rsCommon.MoveNext
            if rsCommon.EOF THen Exit Do
        Loop
    End If

    Call closeRecordset(rsCommon)
    Call ConnClose(Conn)
%>
                                    </table>
                                </td>
                            </tr>
                            <tfoot>
                                <tr>
                                    <td colspan="2">
                                        <input type=submit value="Save Messages" class="button fit" />
                                    </td>
                                </tr>
                            </tfoot>
                        </tbody>
                    </table>
                </form>
            </div>
        </div>
    </div>
</div>
<!-- #include file="../includes/footer.asp"-->