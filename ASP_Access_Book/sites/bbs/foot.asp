<table width="760" border="0" align="center">
  <tr> 
    <td align="center" background="b.jpg"> 
      <% If not (Time2 - Time1)=0 Then%>
      页面执行时间：<font color="#FF3333"> <%= (Time2 - Time1 )*1000 %></font> 毫秒 
      <% End If %>
    </td>
  </tr>
</table>
</body>
</html>