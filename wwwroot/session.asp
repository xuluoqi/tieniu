<!--#include file="connstr.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from yuangong where username='"&session("userid")&"' and password='"&session("password")&"'",conn,1,1
if rs.eof and rs.bof then
response.write "<SCRIPT language=JavaScript>alert('�Բ�����û��Ȩ�޽����ҳ��!');"
response.write"this.location.href='index.asp';</SCRIPT>" 
response.end
end if
rs.close
set rs=nothing
if  session("level")="" then
response.write "<SCRIPT language=JavaScript>alert('�Բ�����û��Ȩ�޽����ҳ��!');"
response.write"this.location.href='index.asp';</SCRIPT>" 
response.end
end if
session.Timeout=60
%>


