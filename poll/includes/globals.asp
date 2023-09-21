<%
	Dim Conn
	Dim ConnStr
	Dim rsCommon
	Dim strSQL
	Dim fso
	Dim total
	Dim strPollName
	Dim strVersion
	Dim show
	Dim msg
	Dim strDBPath
	Dim strCheckIP
	Dim rsAddIP
	Dim strActiveName
	Dim action 
	Dim intCount
	Dim strDir
	Dim strShowAll
	Dim strUpdateQuery
	Dim strUpdateStatQuery
	Dim intPollID
	Dim intPollLength
	Dim strPollQuestion
	Dim blnHideResults
	Dim blnPollRevote
	Dim blnActive
	Dim blnVoted
	Dim strRevote
	Dim strHideResults
	Dim strActive
	Dim datStartDate

	action = Trim(Request.QueryString("action"))
	show = Trim(Request.QueryString("show"))
	msg = Trim(Request.QueryString("msg"))
	
	Response.ExpiresAbsolute = Now() - 2
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "No-Store"
	Response.AddHeader "If-Modified-Since",now
	Response.AddHeader "Last-Modified",now
	Response.Expires = 0
%>
