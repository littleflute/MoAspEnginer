<%@ LANGUAGE= VBScript %>
<html>
<head>
<title>HI!</title>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<% 
If Time >= #12:00:00 AM# And Time < #12:00:00 PM#  Then 
  Greeting = "�糿�ã�" 
Else 
  Greeting = "������!" 
End If
%> 
<%= Greeting %>
</body>
</html>
