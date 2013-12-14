<%@Language="VBSCRIPT"%>
<%
%><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from yuangong where username='"&session("userid")&"' and password='"&session("password")&"'",conn,1,1
if rs.eof and rs.bof then
response.write "<SCRIPT language=JavaScript>alert('对不起，你没有权限进入该页面!');"
response.write"this.location.href='index.asp';</SCRIPT>"
On Error GoTo 0
Err.Raise 9999
end if
rs.close
set rs=nothing
if  session("level")="" then
response.write "<SCRIPT language=JavaScript>alert('对不起，你没有权限进入该页面!');"
response.write"this.location.href='index.asp';</SCRIPT>"
On Error GoTo 0
Err.Raise 9999
end if
session.Timeout=60
%><html><title>员工管理</title><link href="images/hs.css" rel="stylesheet" type="text/css"><style type="text/css"><!--
--></style><link href="admin/zxcss.css" rel="stylesheet" type="text/css"><style type="text/css"><!--
.style1 {
	color: #0033FF;
	font-size: 12px;
}
--></style></head><body><%
set rs2=server.CreateObject("adodb.recordset")
sql="select * from sonjixiang where companyname='"&request("companyname")&"' and peple='"&session("username")&"' and jixiang='"&request("jixiang")&"' "
rs2.open sql,connstr,1,1
%><p align="center" class="style1">&nbsp;</p><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0" class="df"><tr><td><table width="544" height="84" border="0" align="center" cellpadding="0" cellspacing="1" bordercolor="#A4B6D7" bgcolor="#A4B6D7" class="df"><tr bgcolor="#FFCC66"><td width="543" height="30">业务项目：<%= rs2("jixiang") %></td></tr><tr bgcolor="#ECF5FF"><td height="26" bgcolor="#ECF5FF"><table width="544" height="26" border="0" align="center" cellpadding="0" cellspacing="0" class="df"><tr><td height="26" valign="top"><table width="100%" height="26" border="0" align="center" cellpadding="0" cellspacing="0"><tr><td width="13%" height="26" valign="top">项目说明：</td><td width="87%" valign="top"><%= rs2("shuoming") %></td></tr></table></td></tr></table></td></tr><tr bgcolor="#FFCC66"><td height="24"><div align="left"></div><div align="right"></div><div align="left"></div><div align="left"><table width="100%" height="20" border="0" cellpadding="0" cellspacing="0"><tr><td><div align="center"><%
set cn2=conn.execute("select count(*) as mach from shejixiadan where companyname='"&rs2("companyname")&"'")
if cn2("mach")>=1 then
response.Write "已合作"
else
response.Write "未合作"
end if
%></div></td></tr></table></div></td></tr></table></td></tr><%
set rs=server.createobject("adodb.recordset")
sql="select * from richeng where company='"&request("companyname")&"' and peplename='"&session("username")&"' and jixiang='"&request("jixiang")&"' "
rs.open sql,connstr,1,1
%><tr><td height="44"><div align="center"><%
if  rs.eof then
response.Write "<font color=red>"&"没有相关日程!"&"</font>"
else
response.Write "相关日程"
DO While not rs.eof
%></div><table width="544" height="49" border="0" align="center" cellpadding="0" cellspacing="1" bordercolor="#A4B6D7" bgcolor="#A4B6D7" class="df"><tr bgcolor="#ECF5FF"><td width="438" height="24" bgcolor="#FFFFCC"><%= rs("title") %></td><td width="90" bgcolor="#FFFFCC"><%= rs("times") %></td></tr><tr bgcolor="#ECF5FF"><td height="22" colspan="2" valign="top" bgcolor="#ECF5FF"><table width="100%" height="41" border="0" align="center" cellpadding="0" cellspacing="0" class="df"><tr><td height="41" valign="top"><%= rs("shuoming") %></td></tr></table><div align="center"></div></td></tr></table><div align="center"><%
rs.movenext
loop
rs.close
set rs=nothing
end if
%></div></td></tr><tr><td height="39"><div align="center"><a href=javascript:history.go(-1)><a href="javascript:window.close()">关闭</a></a></div></td></tr></table><%
rs2.close
set rs2=nothing
%></body></html>