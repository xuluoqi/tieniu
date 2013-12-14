<!--#include file="connstr.asp"-->
<!--#include file="inc/function.asp"-->
<%if session("level")="" then
response.Write "<script>alert('对不起，你没有权限进入该页面！');history.go(-1)</script>"
Response.End
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="zxcss.css" rel="stylesheet" type="text/css">
<link href="admin/zxcss.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style3 {color: #FF0000}
.style7 {
	color: #3366FF;
	font-size: 10pt;
}
.style8 {color: #D4D0C8}
-->
</style>
<script language="javascript" src="inc/func.js" type="text/javascript"></script>
</head>

<body>
<p>
  <%
   set rs=server.createobject("adodb.recordset")
sql="select * from jishiben where 1=1"
if not CheckUserPermission("733") then sql=sql&" and yewuyuan='"&session("username")&"'"
sql=sql&" order by times desc "
rs.open sql,conn,1,1 
page=cint(request("page"))
rs.pagesize=25
count=rs.pagesize
pgnm=rs.pagecount
if page="" or clng(page)<1 then page=1
if clng(page)>pgnm then page=pgnm
if pgnm>0 then rs.absolutepage=page
if request("page1")<>"" then response.Redirect "jishiben.asp?page="&request("page1")
%>
</p>
<table width="79%" height="96"  border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFCC" bgcolor="#FFFFFF">
  <tr>
    <td width="100%" height="18"><div align="right"><span class="style7"><a href="jishiben_add.asp">添加日记&nbsp;&nbsp;&nbsp;</a></span></div></td>
  </tr>
  <tr>
    <td height="78" bgcolor="#FFFFFF"><table width="100%" height="64" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
          <td bordercolor="#FFFFCC">
		  <table width="100%" height="18" border="0" align="center" cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#CCCCCC">
              <% DO While not rs.eof %>
              <tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#efefef'" onMouseOut="this.bgColor='#ffffff'">
                <td height="18"><span class="style8">　
                <a href="javascript:openEditScript('jishiben_list.asp?id=<%=rs("id")%>',750,370)"><%=rs("title")%></a></span></td>
                <td width="100">&nbsp;<%=rs("yewuyuan")%></td>
                <td width="170">&nbsp;<%=rs("times")%></td>
                <td width="100"><div align="center" class="style8">
                    <%if year(rs("times"))=year(now) and month(rs("times"))=month(now) and day(rs("times"))=day(now) then %>
                    <a href="jishiben_xiugai.asp?id=<%=rs("id")%>">修改</a>
                    <%else
					response.Write "&nbsp;"
					end if%>
                </div></td>
              </tr>
              <% 
			  rs.movenext 
	   i=i+1
	   if i>=count then exit do 
	   loop
	   %>
          </table></td>
        </tr>
        <tr><form action="" method="post" name="form1">
	            <td height="24" align="center" bordercolor="#FFFFCC" bgcolor="#FFFFFF">共有记录&nbsp;<span class="style3"><%= rs.recordcount %>&nbsp;</span>个,当前页<span class="style3"> <%= page %>/<%= pgnm %></span>：
                  <% If page>1 Then %>
                  <a href="jishiben.asp?page=<%= page-1 %>">上一页</a>
                  <% Else %>
                  <%= "上一页" %>
                  <% End If %>
                  <% If page<>pgnm Then %>
                  <a href="jishiben.asp?page=<%=page+1 %>"> 下一页</a>
                  <% Else %>
                  <%= "下一页" %>
                  <% End If %>
            <input name="page1" type="text" id="page1" size="1">                  <input type="submit" name="Submit" value="go"></td>
		  </form>
        </tr>
    </table></td>
  </tr>
</table>
<%rs.close
	   set rs=nothing%>
</body>
</html>

