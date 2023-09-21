<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
  Response.Cookies("EZPollAdmin").Expires = Date -1  
  Session.Abandon()

  Response.Redirect "admin_login.asp"
%>