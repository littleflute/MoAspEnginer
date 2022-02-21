<%@ Language=VBScript %>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="guest_sub.asp" -->
<%if Request.Form("from") = "Delete" then 
        gbPassWord=request.form("password")
        if gbPassWord<>PassWord then
             Msg("<font color=#000000><li>无效密码，请核对后再试。</li></font>")
             Response.End 
        end if
        Response.Cookies("UserInfo")("guestPassWord")=trim(request.form("password"))
        Response.Cookies("User").Expires=  dateAdd("d", 1, now)
    	dim gbID
	    gbID=request("ID")
	    Set book=server.createobject("adodb.recordset")
        SQLstr="Delete * from guest where ID = " & gbID
        conn.Execute SQLstr
        set book=nothing
        conn.close
        set conn=nothing
	 Redirect "guest_list.asp?PageNo=" & Request("PageNo") & "","您已成功的删除了该留言,2秒之后系统自动返回."
end if
%>	