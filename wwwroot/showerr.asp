<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
dim sysstatu,rssys,dqflag
set rssys = server.createobject("adodb.recordset")
rssys.open "select * from sysconfig",conn,1,1
if not rssys.eof then
sysstatu = rssys("SystemStatu")
ExMaxNumDate = rssys("ExMaxNumDate")
end if
rssys.close
set rssys = nothing
dqflag = false
if sysstatu=0 and (isnull(ExMaxNumDate) or ExMaxNumDate>=date()) then response.redirect "index.asp"
%><html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312" /><title>ϵͳ��ͣ����[��ɴӰ¥�������]</title><style type="text/css"><!--
* {
	font-size: 12px;
	color: #000000;
}
--></style></head><body><table width="100%" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC"><tr><th height="30" bgcolor="#E2DEE2">������ԭ��ϵͳ��ͣʹ��</th></tr><tr><td width="100%" bgcolor="#FFFFFF" style="padding:10px 0 10px 3px"><%
if sysstatu=1 then
response.write conn.execute("select StopReadme from sysconfig")(0)
elseif not isnull(ExMaxNumDate) and ExMaxNumDate<date() then
dqflag = true
response.write "���ʹ���ѵ��ڣ���ͣʹ�á�"
end if
%></td></tr><%
if not dqflag then
%><tr align="middle"><td width="100%" height="20" align="center" bgcolor="#FFFFFF"><a href="admin_login.asp" target="_blank">ϵͳ����</a></td></tr><%
end if
%></table></body></html>