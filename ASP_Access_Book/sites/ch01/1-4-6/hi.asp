<%@ LANGUAGE= VBScript %>
<html>
<head>
<title>HI!</title>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<% 
If Time >= #12:00:00 AM# And Time < #12:00:00 PM#  Then 
  Greeting = "Ôç³¿ºÃ£¡" 
Else 
  Greeting = "³ÔÁËÂð!" 
End If
%> 
<%= Greeting %>
</body>
</html>
