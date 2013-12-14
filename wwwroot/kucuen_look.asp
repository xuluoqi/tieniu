<%@Language="VBSCRIPT"%>
<%
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
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from yunyong where [type]=1 and ishidden=0 order by px asc ",conn,1,1
%><link href="admin/zxcss.css" rel="stylesheet" type="text/css"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC"><tr bgcolor="#99FFFF"><td width="41%"><div align="center">类别</div></td><td width="36%"><div align="center">数量</div></td></tr><%
while not rs.eof
%><tr bgcolor="#FFFFFF"><td><div align="left"></div><div align="center"><%= rs("yunyong") %></div></td><td><div align="left"></div><div align="center"><%= rs("sl") %></div></td></tr><%
rs.movenext
wend
rs.close
set rs=nothing
%></table>