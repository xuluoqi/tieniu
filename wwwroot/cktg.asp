<%@Language="VBSCRIPT"%>
<%
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"><html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>员工管理</title><link href="images/hs.css" rel="stylesheet" type="text/css"><style type="text/css"><!--
.style3 {color: #FF0000}
.df {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-style: normal;
	line-height: normal;
	font-weight: normal;
	font-variant: normal;
}
--></style><%
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
%><script language="javascript">
function openEditScript(url, width, height){
	var Win = window.open(url,"openEditScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=no' );
}
</script><link href="zxcss.css" rel="stylesheet" type="text/css"><style type="text/css"><!--
.style4 {
	color: #3333FF;
	font-size: 10px;
}
--></style><link href="admin/zxcss.css" rel="stylesheet" type="text/css"><style type="text/css"><!--
.style5 {font-size: 14px}
--></style></head><body><p align="center"><%
if request("DeleId")<>""then
conn.execute("delete*from tonggao where id="&request("DeleId")&"")
response.Write "<script>alert('删除完成！');location='tonggao.asp'</script>"
end if
set rs=server.createobject("adodb.recordset")
sql="select * from tonggao order by  times desc"
rs.open sql,connstr,1,1
page=cint(request("page"))
rs.pagesize=15
count=rs.pagesize
pgnm=rs.pagecount
if page="" or clng(page)<1 then page=1
if clng(page)>pgnm then page=pgnm
if pgnm>0 then rs.absolutepage=page
%><span class="style4 style5">通告列表 </span></p><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0" class="df"><tr><td height="56"><table width="661" height="53" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFCCFF" bgcolor="#FFFFFF"><tr bgcolor="#ECF5FF"><td height="24" bgcolor="#FFFFFF"><table width="657" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC"><%
DO While not rs.eof
%><tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#efefef'" onMouseOut="this.bgColor='#ffffff'"><td width="496" height="19"> 　<a href="javascript:openEditScript('cktg2.asp?id=<%= rs("id") %>',780,370)"> <%= rs("title") %></a></td><td width="158">&nbsp;&nbsp;<%= rs("times") %><div align="center"></div></td></tr><%
rs.movenext
i=i+1
if i>=count then exit do
loop
rs.close
set rs=nothing
%></table></td></tr><tr bgcolor="#ECF5FF"><td height="24" bgcolor="#FFFFFF"><div align="center">当前页<span class="style3"><%=  page %>/<%=  pgnm %></span>：
              <%
If page>1 Then
%><a href="tonggao.asp?page=<%=  page-1 %>">上一页</a><%
Else
%><%=  "上一页" %><%
End If
If page<>pgnm Then
%><a href="tonggao.asp?page=<%= page+1 %>"> 下一页</a><%
Else
%><%=  "下一页" %><%
End If
%></div></td></tr></table></td></tr></table><p>&nbsp;</p></body></html>