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
%><html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>无标题文档</title></head><%
if request("id")<>"" then
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from richeng where id="&request("id")&"",conn,1,1
if rs("times")<>date then response.Write "<script>alert('你没有权限删除数据!');history.go(-1)</script>"
set rs2=server.CreateObject("adodb.recordset")
rs2.open "select * from richeng where company='"&rs("company")&"'and jixiang='"&rs("jixiang")&"' order by times desc",conn,1,1
if rs2.recordcount<=1 then
conn.execute("delete from sonjixiang where companyname='"&rs("company")&"' and jixiang='"&rs("jixiang")&"'")
else
if int(rs2("id"))=int(request("id")) then
rs2.absoluteposition=2
conn.execute("update sonjixiang set richeng_time=#"&rs2("times")&"# where companyname='"&rs("company")&"' and jixiang='"&rs("jixiang")&"'")
end if
end if
conn.execute("delete from  richeng where id="&request("id"))
rs.close
set rs=nothing
rs2.close
set rs2=nothing
conn.close
set conn=nothing
response.Write "<script>alert('删除成功!');location='right.asp'</script>"
end if
%><body></body></html>