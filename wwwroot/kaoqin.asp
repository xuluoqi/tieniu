<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
%><html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>无标题文档</title><link href="admin/zxcss.css" rel="stylesheet" type="text/css"><script src="Js/Calendar.js"></script><link href="Css/calendar-blue.css" rel="stylesheet"><style type="text/css"><!--
.style3 {color: #FF0000}
.style4 {color: #FFCC99}
--></style></head><body><%
if session("username")="" then
response.Write "<script>alert('对不起,你没权限进入该页面！');history.go(-1)</script>"
On Error GoTo 0
Err.Raise 9999
end if
%><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#efefef"><form action="" method="post" name="form1"><tr><td width="90%" bgcolor="#FFFFFF"><div align="left">&nbsp;&nbsp;按时间查询:
        <input name="fromtime" type="text" id="fromtime" size="10" value="<%
if  request("fromtime")="" then
response.Write date-30
else
response.Write request("fromtime")
end if
%>">
&nbsp;&nbsp; <span class="font"><A onclick="return showCalendar('fromtime', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></A></span>&nbsp;到：
<input name="totime" type="text" id="totime" size="10" value="<%
if request("totime")="" then
response.Write date()
else
response.Write request("totime")
end if
%>">
&nbsp;&nbsp;&nbsp;<span class="font"><A onclick="return showCalendar('totime', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></A></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input name="chaxun1" type="submit" id="chaxun1" value="查询">
&nbsp;</div></td><td width="10%" bgcolor="#FFFFFF">&nbsp;<a href="#" onClick="javascript:history.go(-1)">返回</a></td></tr></form></table><br><%
set rs=server.CreateObject("adodb.recordset")
if request("fromtime")<>"" and request("totime")<>"" and request("peplename")="" then
rs.open "select * from kaoqing where times>=#"&request("fromtime")&"# and times<=#"&request("totime")&"# and peplename='"&session("username")&"' order by times desc",conn,1,1
else
rs.open "select * from kaoqing where peplename='"&session("username")&"' order by times desc",conn,1,1
end if
page=cint(request("page"))
rs.pagesize=25
count=rs.pagesize
pgnm=rs.pagecount
if page="" or clng(page)<1 then page=1
if clng(page)>pgnm then page=pgnm
if pgnm>0 then rs.absolutepage=page
%><table width="98%" height="0"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC"><tr bgcolor="#FFFFFF"><td width="22%"><div align="center"></div></td><td colspan="2"><div align="center">早上</div></td><td colspan="2"><div align="center">下午</div></td><td colspan="2"><div align="center">晚上</div></td></tr><tr bgcolor="#FFFFFF"><td height="14"><div align="center">日期</div></td><td width="14%"><div align="center"></div><div align="center">上班</div></td><td width="13%"><div align="center">下班</div></td><td width="13%"><div align="center">上班</div></td><td width="13%"><div align="center">下班</div></td><td width="13%"><div align="center">上班</div></td><td width="12%"><div align="center">下班</div></td></tr><%
do while not rs.eof
%><tr bgcolor="#FFFFFF"><td>&nbsp;&nbsp;<%= rs("times") %>
	[
	  <%
select case WEEKDAY(rs("times"))
case 1
response.Write "星期日"
case 2
response.Write "星期一"
case 3
response.Write "星期二"
case 4
response.Write "星期三"
case 5
response.Write "星期四"
case 6
response.Write "星期五"
case 7
response.Write "星期六"
end select
%>
    ]    </td><td><div align="center"></div><div align="center"><%
if isnull(rs("time1")) then
response.Write "无"
else
response.Write hour(rs("time1"))&":"&minute(rs("time1"))
end if
%></div></td><td><div align="center"><%
if isnull(rs("time2")) then
response.Write "无"
else
response.Write hour(rs("time2"))&":"&minute(rs("time2"))
end if
%></div></td><td><div align="center"><%
if isnull(rs("time3")) then
response.Write "无"
else
response.Write hour(rs("time3"))&":"&minute(rs("time3"))
end if
%></div></td><td><div align="center"><%
if isnull(rs("time4")) then
response.Write "无"
else
response.Write hour(rs("time4"))&":"&minute(rs("time4"))
end if
%></div></td><td><div align="center"><%
if isnull(rs("time5")) then
response.Write "无"
else
response.Write hour(rs("time5"))&":"&minute(rs("time5"))
end if
%></div></td><td><div align="center"><%
if isnull(rs("time6")) then
response.Write "无"
else
response.Write hour(rs("time6"))&":"&minute(rs("time6"))
end if
%></div></td></tr><%
rs.movenext
i=i+1
if i>=count then exit do
loop
%></table><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0"><tr><td><div align="center">共有记录<span class="style3"><%=  rs.recordcount %>&nbsp;</span>个,当前页<span class="style3"><%=  page %>/<%=  pgnm %></span>：
        <%
If page>1 Then
%><a href="kaoqin.asp?page=<%=  page-1 %>&fromtime=<%= request("fromtime") %>&totime=<%= request("totime") %>">上一页</a><%
Else
%><%=  "上一页" %><%
End If
If page<>pgnm Then
%><a href="kaoqin.asp?page=<%= page+1 %>&fromtime=<%= request("fromtime") %>&totime=<%= request("totime") %>"> 下一页</a><%
Else
%><%=  "下一页" %><%
End If
%></div></td></tr></table><%
rs.close
set rs=nothing
%></body></html>