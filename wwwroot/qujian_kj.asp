<%@Language="VBSCRIPT"%>
<%
%><html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>日报表打印</title></head><frameset rows="25,*" cols="*" framespacing="2" frameborder="NO" border="2"><frame src="print_top.asp?id=<%= request("id") %>" name="top" scrolling="NO" noresize id="top" >
  <frame src="qujian.asp?id=<%= request("id") %>" name="main11" id="main11">
</frameset><noframes><body></body></noframes></html>