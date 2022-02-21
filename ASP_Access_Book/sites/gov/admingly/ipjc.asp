<%
   Private Function lockip()
    Dim strIPAddrs
    If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
        strIPAddrs = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
        strIPAddrs = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
        strIPAddrs = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
    Else
        strIPAddrs = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    End If
    lockip = Trim(Mid(strIPAddrs, 1, 30))
End Function

    dim yhip,kill,badip
	    kill=false
	yhip=lockip()
	sql="select badip from killip order by times desc"
	set rs=conn.execute(sql)
	  if not rs.eof then
	  badip=rs.getrows
	  for i=0 to ubound(badip,2)
	    if yhip=badip(0,i) then
		  kill=true
		  exit for
		  end if
		  next 
   if kill=true then
   response.write"<script>alert('对不起,您的IP地址被锁定了,不能登陆');window.location.href='error.htm';</script>"
       end if
	   end if
%>
