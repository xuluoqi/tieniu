<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from yuangong where username='"&session("userid")&"' and password='"&session("password")&"'",conn,1,1
if rs.eof and rs.bof then
response.write "<SCRIPT language=JavaScript>alert('�Բ�����û��Ȩ�޽����ҳ��!');"
response.write"this.location.href='index.asp';</SCRIPT>"
On Error GoTo 0
Err.Raise 9999
end if
rs.close
set rs=nothing
if  session("level")="" then
response.write "<SCRIPT language=JavaScript>alert('�Բ�����û��Ȩ�޽����ҳ��!');"
response.write"this.location.href='index.asp';</SCRIPT>"
On Error GoTo 0
Err.Raise 9999
end if
session.Timeout=60
%><html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>�ޱ����ĵ�</title><link href="../admin/zxcss.css" rel="stylesheet" type="text/css"><style type="text/css"><!--
.style1 {color: #0000FF}
.style2 {color: #6666FF}
--></style><link href="admin/zxcss.css" rel="stylesheet" type="text/css"></head><body><table width="774" height="108" border="0" align="center" cellpadding="0 " cellspacing="0" bgcolor="#66FFFF"><tr><td height="22" colspan="5" bgcolor="#FFCC33"><div align="center" class="style1">���ҵ�������</div></td></tr><tr bgcolor="#CCCCFF"><td width="228" height="18"><div align="center">��Ŀ��˾</div></td><td width="188"><div align="center">��Ŀ</div></td><td width="134"><div align="center">ҵ��</div></td><td width="138"><div align="center">���</div></td><td width="86">�µ�ʱ��</td></tr><%
set rs=server.CreateObject("adodb.recordset")
sql="select * from shejixiadan where yewuyuan='"&session("username")&"' order by times desc"
rs.open sql,connstr,1,1
set cn=conn.execute("select sum(yewuchoucheng)as choucheng from shejixiadan where yewuyuan='"&session("username")&"' and yewuyichoucheng="&true&"")
set cn1=conn.execute("select sum(yewuchoucheng)as choucheng1 from shejixiadan where yewuyuan='"&session("username")&"' and yewuyichoucheng="&false&"")
set cn2=conn.execute("select sum(yewuchoucheng)as choucheng2 from shejixiadan where yewuyuan='"&session("username")&"'")
set cn3=conn.execute("select sum(feiyong)as yeji from shejixiadan where yewuyuan='"&session("username")&"'")
while not rs.eof
if rs("yewuyichoucheng")=true then
%><tr><td height="20"><div align="center"><%= rs("companyname") %></div></td><td><div align="center"><%= rs("jixiang") %></div></td><td><div align="center"><%= rs("feiyong") %><span class="style2">Ԫ</span></div></td><td><div align="center"><%= rs("yewuchoucheng") %>Ԫ</div></td><td><%= rs("times") %></td></tr><%
else
%><tr><td height="20"><div align="center"><%= rs("companyname") %></div></td><td><div align="center"><%= rs("jixiang") %></div></td><td><div align="center"><%= rs("feiyong") %><span class="style2">Ԫ</span></div></td><td><div align="center"><%= rs("yewuchoucheng") %>Ԫ��<font color="blue">δ���</font>��</div></td><td><%= rs("times") %></td></tr><%
end if
rs.movenext
wend
rs.close
set rs=nothing
%><tr bgcolor="#FFCCCC"><td height="20"><div align="center" class="style2">�ѳ�ɺϼƣ�<%= cn("choucheng") %>Ԫ</div></td><td><div align="center" class="style2">δ��ɺϼƣ�<%= cn1("choucheng1") %>Ԫ</div></td><td><div align="center"><span class="style2">��ҵ��:<%= cn3("yeji") %>Ԫ</span></div></td><td><div align="center" class="style2"><div align="left">�ܳ�ɣ�<%= cn2("choucheng2") %>Ԫ</div></div></td><td>&nbsp;</td></tr></table><p>&nbsp;</p><%
conn.close
set conn=nothing
%></body></html>