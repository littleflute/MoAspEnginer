<%@ Language=VBScript %>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="guest_sub.asp" -->
<%if Request.Form("from") = "Delete" then 
        gbPassWord=request.form("password")
        if gbPassWord<>PassWord then
             Msg("<font color=#000000><li>��Ч���룬��˶Ժ����ԡ�</li></font>")
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
	 Redirect "guest_list.asp?PageNo=" & Request("PageNo") & "","���ѳɹ���ɾ���˸�����,2��֮��ϵͳ�Զ�����."
end if
%>	