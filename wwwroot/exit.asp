<%@Language="VBSCRIPT"%>
<%
session("username") =""
session("level")=""
session("userid")=""
session.abandon
response.redirect "index.asp"
%>