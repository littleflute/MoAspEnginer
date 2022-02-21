<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
Dim rs
Dim rs_numRows
Set rs = Server.CreateObject("ADODB.Recordset")
if Trim(Request.QueryString("order"))="2" then
rs.Source = "SELECT * FROM uers order by totle desc,time desc"
else
rs.Source = "SELECT * FROM uers order by time desc"
end if
rs.Open rs.Source,conn,1,1
rs_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_numRows = rs_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rs_total
Dim rs_first
Dim rs_last

' set the record count
rs_total = rs.RecordCount

' set the number of rows displayed on this page
If (rs_numRows < 0) Then
  rs_numRows = rs_total
Elseif (rs_numRows = 0) Then
  rs_numRows = 1
End If

' set the first and last displayed record
rs_first = 1
rs_last  = rs_first + rs_numRows - 1

' if we have the correct record count, check the other stats
If (rs_total <> -1) Then
  If (rs_first > rs_total) Then
    rs_first = rs_total
  End If
  If (rs_last > rs_total) Then
    rs_last = rs_total
  End If
  If (rs_numRows > rs_total) Then
    rs_numRows = rs_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rs_total = -1) Then

  ' count the total records by iterating through the recordset
  rs_total=0
  While (Not rs.EOF)
    rs_total = rs_total + 1
    rs.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rs.CursorType > 0) Then
    rs.MoveFirst
  Else
    rs.Requery
  End If

  ' set the number of rows displayed on this page
  If (rs_numRows < 0 Or rs_numRows > rs_total) Then
    rs_numRows = rs_total
  End If

  ' set the first and last displayed record
  rs_first = 1
  rs_last = rs_first + rs_numRows - 1
  
  If (rs_first > rs_total) Then
    rs_first = rs_total
  End If
  If (rs_last > rs_total) Then
    rs_last = rs_total
  End If

End If
%>

<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rs
MM_rsCount   = rs_total
MM_size      = rs_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rs_first = MM_offset + 1
rs_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rs_first > MM_rsCount) Then
    rs_first = MM_rsCount
  End If
  If (rs_last > MM_rsCount) Then
    rs_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<html>
<head>
<title>用户列表</title>
<!-- #include file="an.htm" -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="475" height="24"><a href="index.asp">论坛论坛</a>→<% if Trim(Request.QueryString("order"))="2" then %> 发贴排名 <% Else %>
用户列表  <% End If %>
</td>
    <td width="283">|<a href="uerlist.asp">用户列表</a> | <a href="adduer.asp">用户注册</a>| 
      <a href="find.asp">查找会员</a>| <a href="uerlist.asp?order=2">发贴排名</a>|</td>
  </tr>
    <tr align="center"> 
    <td height="15" colspan="9" background="pic/h3.gif" class="p9"><strong><font color="#FFFFFF"><% if Trim(Request.QueryString("order"))="2" then %> 发贴排名 <% Else %>
用户列表<% End If %></font></strong></td>
  </tr>
</table>
<table width="760" height="31%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#83BB37"  bordercolordark="#83BB37" bordercolorlight="#FFFFFF">

  <tr> 
    <td height="16" colspan="9" class="p9"> 　主要干部｜秘书处｜宣传部｜公关部｜OEC活动部｜</td>
  </tr>
  <tr align="center" bgcolor="#e3f1d1" class="pbutton"> 
    <td width="10%" height="20" class="p9">论坛ID</td>
    <td width="8%" class="p9">姓名</td>
    <td width="20%" class="p9">E-mail</td>
    <td width="9%" class="p9">Oicq</td>
    <td width="10%" class="p9">电话</td>
    <td width="12%" class="p9">住址</td>
    <td width="8%" class="p9">发文章数</td>
    <td width="8%" class="p9">等级</td>
    <td width="18%" class="p9">最后登陆</td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rs.EOF)) 
%>
  <tr align="center"> 
    <td height="20" bgcolor="#FFFFFF" class="p9"><A HREF="uerparticle.asp?<%= MM_keepURL & MM_joinChar(MM_keepURL) & "id=" & rs.Fields.Item("id").Value %>"><%=(rs.Fields.Item("id").Value)%></A></td>
    <td height="20" bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("name").Value)%></td>
    <td height="20" bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("email").Value)%></td>
    <td bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("qq").Value)%></td>
    <td height="20" bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("tel").Value)%></td>
    <td height="20" bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("addr").Value)%></td>
    <td bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("totle").Value)%></td>
    <td bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("dengji").Value)%></td>
    <td bgcolor="#FFFFFF" class="p9"><% if rs("lasttime")<>"" then %><%=(rs.Fields.Item("lasttime").Value)%><% Else %><%=(rs.Fields.Item("time").Value)%><% End If %>
</td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs.MoveNext()
Wend
%>
  <tr> 
    <td colspan="9" bgcolor="#e3f1d1" class="p9" > <% If MM_offset <> 0 Then %> <A HREF="<%=MM_moveFirst%>">第一页</A><% Else %>第一页<% End If ' end MM_offset <> 0 %>
      　　 
      <% If MM_offset <> 0 Then %> <A HREF="<%=MM_movePrev%>">上一页</A><% Else %>
 上一页<% End If ' end MM_offset <> 0 %>
      　　 
      <% If Not MM_atTotal Then %> <A HREF="<%=MM_moveNext%>">下一页</A> <% Else %>下一页
<% End If ' end Not MM_atTotal %>
      　 
      <% If Not MM_atTotal Then %> <A HREF="<%=MM_moveLast%>">最后一页</A> <% Else %>最后一页
<% End If ' end Not MM_atTotal %>
      　共有<%=(rs_total)%>名会员</td>
  </tr>
  <tr> 
    <td colspan="9" align="center" bordercolor="#CCFF99" > <form name="form1" method="post" action="find2.asp">
        <hr>
        会员查找： 填写： 
        <input name="name" type="text" id="name3">
        选择条件： 
        <select name="select">
          <option value="1">姓 名</option>
          <option value="2">论坛ID</option>
        </select>
        <input type="submit" name="Submit" value="　搜　索　">
      </form></td>
  </tr>
</table>
<%
rs.Close()
Set rs = Nothing
%>
<!--#include file="foot.asp"-->

