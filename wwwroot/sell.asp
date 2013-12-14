<!--#include file="connstr.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="session.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="admin/zxcss.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style3 {color: #FF0000}
-->
</style>
<script language="javascript" src="inc/func.js" type="text/javascript"></script>
<script language="javascript">
function check()
{
if(!CheckIsNull(document.Form1.shejishi,"请选择摄影师！")) return false;
}
</script>
</head>

<body>
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sell_type",conn,1,1
response.write"<script language = ""JavaScript"">"
response.write"var onecount;"
response.write"onecount=0;"
response.write"subcat = new Array();"
count=1
while not rs.eof 
response.write"subcat["&count&"] = new Array('"& rs("dj")&"','"&rs("id")&"');"
count = count + 1
rs.movenext 
wend 
rs.close
set rs=nothing
response.write"onecount="&count&";"
response.write"function change(month)"
response.write"{"
response.write"var month=month;"
response.write"var i=1;"
response.write"var kk='';"
response.write"var jj='';"
response.write"for (i=1;i < onecount; i++)"
response.write"{"
response.write"if (subcat[i][1] == month)"
response.write"{"
response.write"kk+=subcat[i][0];"
response.write"}"
response.write"}"
response.Write "document.Form1.dj.value=kk ;"
response.write"}"
response.Write "</script>"
select case request("action")
case "added2"
if request("name")="" then
response.Write "<script> alert('请填写类别名称！');history.go(-1)</script>"
Response.End
end if
if not isnumeric(request("dj")) then
response.Write "<script> alert('单价金额填写错误,只能是数字！');history.go(-1)</script>"
Response.End
end if
if not isnumeric(request("sl")) then
response.Write "<script> alert('数量填写错误,只能是数字！');history.go(-1)</script>"
Response.End
end if
conn.execute("insert into sell_jilu (yuangong_id,xiangmu_id,[name],dj,sl,beizhu,times)  values ("&Request("shejishi")&","&request("id")&",'"&conn.execute("select [name] from sell_type where  id="&request("name")&"")(0)&"',"&request("dj")&","&request("sl")&",'"&htmlencode2(request("beizhu"))&"',now())")
response.Write "<script>alert('添加记录成功！');location='fujia.asp?id="&request("id")&"'</script>"
Response.End
case "edited"
name11=split(request("name"),", ")
id11=split(request("id"),", ")
dj11=split(request("dj"),", ")
for i=lbound(id11) to ubound(id11)
if name11(i)="" or not isnumeric(dj11(i)) then
response.Write "<script>alert('错误请检查，不能为空，单价只能是数字！');history.go(-1)</script>"
Response.End
end if
conn.execute("update sell_type set [name]='"&name11(i)&"',dj="&dj11(i)&" where id="&id11(i)&"")
next
response.Write "<script>alert('修改成功！');location='sell.asp?action=type&id2="&request("id2")&"'</script>"
response.End
case "added"
if request("name")="" then
response.Write "<script> alert('请填写类别名称！');history.go(-1)</script>"
Response.End
end if
if not isnumeric(request("dj")) then
response.Write "<script> alert('物品单价填写错误，请检查！');history.go(-1)</script>"
Response.End
end if
conn.execute("insert into sell_type (name,dj) values ('"&request("name")&"',"&request("dj")&")")
response.Write "<script>alert('添加成功！');location='sell.asp?action=type&id2="&request("id2")&"'</script>"
Response.End
case "dele"
conn.execute("delete from sell_type where id="&request("id")&"")
response.Write "<script>location='sell.asp?action=type&id2="&request("id2")&"'</script>"
case "type" %>
<br>
<br>
<table width="526" height="30" border="0" align="center" cellpadding="0" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#CCCCCC">
<form name="Form1"  method="post" action="sell.asp?action=added">
  <tr>
    <td width="473" height="26" bgcolor="#FFFFFF"><div align="left">&nbsp;&nbsp;添加类别: 
          <input name="name" type="text" id="name" size="13">
&nbsp;单价:
<input name="dj" type="text" id="dj" size="5" onKeyUp="value=value.replace(/[^\d]/g,'')" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
元&nbsp;&nbsp;
<input type="submit" name="Submit" value="添加">
    <input name="id2" type="hidden" id="id2" value="<%=request("id2")%>">
    </td>
    <td width="50" align="center" bgcolor="#FFFFFF"><a href="fujia.asp?id=<%=request("id2")%>" onClick="history.go(-1)">返回</a></td>
  </tr>
  </form>
</table>
<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sell_type",conn,1,1
%>
<table width="526" border="0" align="center" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
<form action="sell.asp?action=edited" method="post" name="Form1">
    <tr bgcolor="#FFFFFF">
      <td height="22" colspan="3" align="center">类别列表</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="226" height="18" align="center">类别</td>
      <td width="148" align="center">单价</td>
      <td width="148" align="center">管理</td>
    </tr>
	<%while not rs.eof%>
    <tr bgcolor="#FFFFFF">
      <td height="18" align="center">
        <input name="name" type="text" id="name" size="20" value="<%=rs("name")%>">
        <input name="id" type="hidden" id="id" value="<%=rs("id")%>">
      </td>
      <td align="center">
        <input name="dj" type="text" id="dj" size="5" value="<%=rs("dj")%>">
      元</td>
      <td align="center"><a href="sell.asp?id=<%=rs("id")%>&id2=<%=request("id2")%>&action=dele" onClick="return confirm('确定要删除吗？')">删除</a></td>
    </tr>
	<%rs.movenext
	wend 
	rs.close
	set rs=nothing%>
    <tr bgcolor="#FFFFFF">
      <td height="30" colspan="3" align="center">
        <input type="submit" name="Submit4" value="修改">
        <input name="id2" type="hidden" id="id2" value="<%=request("id2")%>">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="reset" name="Submit5" value="重置">
      </td>
    </tr>
	
  </form>
</table>
<%case "add"%>
<br>
<br>
<table width="80%"  border="0" align="center" cellpadding="0" cellspacing="0">
<form action="sell.asp?action=added2" method="post" name="Form1" onSubmit="return check()">
  <tr>
    <td align="center"><select name="shejishi" id="shejishi">
            <option value="" selected>请选择摄影师</option>
            <% 
	  Set S_Rs=Conn.Execute("Select distinct userid From xiadan Where type=4 and xiangmu_id="&Request("id"))
	  Do While Not S_Rs.Eof
	  %>
            <option value="<%=Conn.Execute("Select ID from yuangong where username='"&S_Rs("userid")&"'")(0)%>"><%=Conn.Execute("Select peplename from yuangong where username='"&S_Rs("userid")&"'")(0)%></option>
            <%
	  S_Rs.MoveNext
	  Loop
	  S_Rs.Close
	  Set S_Rs=Nothing
	  %>
          </select>
          &nbsp;&nbsp;
          导购物品:
          <select name="name" id="name"  onChange="change(this.options[this.selectedIndex].value)">
            <option value="">请选择</option>
            <%
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from sell_type ",conn,1,1
		while not rs.eof
		%>
            <option value="<%=rs("id")%>"><%=rs("name")%></option>
            <%rs.movenext
		wend
		rs.close
		set rs=nothing%>
          </select>
          &nbsp;&nbsp;单价:
          <input name="dj" type="text" id="dj" size="5" onKeyUp="value=value.replace(/[^\d]/g,'')   "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
          &nbsp;&nbsp; 数量:
          <input name="sl" type="text" id="sl" size="5" onKeyUp="value=value.replace(/[^\d]/g,'')   "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
          
      </td></tr>
  <tr>
    <td align="center">
      <textarea name="beizhu" cols="70" rows="7" id="beizhu"></textarea>
    </td>
  </tr>
  <tr>
    <td height="46" align="center">
      <input type="submit" name="Submit2" value="提交">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <input type="button" name="Submit3" value="返回" onClick="history.go(-1)">
      <input name="id" type="hidden" id="id" value="<%=request("id")%>">    </td>
  </tr>
  </form>
</table>
<br>
<table width="81%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="17%" align="center" valign="top"><div align="right">常用语：</td>
    <td width="83%" align="center">
      <div align="left">
        <textarea name="yongyu" cols="80" rows="7" id="textarea2"><%if not isnull(conn.execute("select yongyu from two_yongyu where userid='"&"admin"&"'")(0)) then 
	  response.Write conn.execute("select yongyu from two_yongyu where userid='"&"admin"&"'")(0)
	  end if%>
      </textarea>
    </td></tr>
</table>
<%case else%>
<br>
<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#efefef">
  <tr bgcolor="#FFFFFF">
    <td width="86%" bgcolor="#efefef">&nbsp;<a href="sell.asp?action=type"><strong>类别管理</strong></a></td>
    <td width="14%" align="center" bgcolor="#efefef"><a href="sell.asp?action=add">添加记录</a></td>
  </tr>
</table>
<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sell_jilu order by times desc",conn,1,1
page=cint(request("page"))
rs.pagesize=3
count=rs.pagesize
pgnm=rs.pagecount
if page="" or clng(page)<1 then page=1
if clng(page)>pgnm then page=pgnm
if pgnm>0 then rs.absolutepage=page
%>
<%if rs.eof then%>
<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>目前还没有数据</td>
  </tr>
</table>
<%else%>
<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
<%do while not rs.eof %>
  <tr bgcolor="#FFFFFF">
    <td width="37%">&nbsp;项目:<%=rs("name")%></td>
    <td width="17%">&nbsp;单价:<%=rs("dj")%>元</td>
    <td width="21%">&nbsp;&nbsp;数量:<%=rs("sl")%></td>
    <td width="25%">&nbsp;时间:<%=rs("times")%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="4">&nbsp;备注:<%=encode(rs("beizhu"))%></td>
  </tr>
  <%rs.movenext
  i=i+1
  if i>=count then exit do
  loop
  %>
</table>
<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">共有记录&nbsp;<font color="#FF0000"><%= rs.recordcount %></font>&nbsp;个　当前页<span class="style3"> <%= page %>/<%= pgnm %></span>：
        <% If page>1 Then %>
          <a href="sell.asp?page=<%= page-1 %>">上一页</a>
        <% Else %>
        <%= "上一页" %>
        <% End If %>
        <% If page<>pgnm Then %>
          <a href="sell.asp?page=<%=page+1 %>"> 下一页</a>
        <% Else %>
        <%= "下一页" %>
        <% End If %>    </td>
  </tr>
</table>
<%rs.close
  set rs=nothing%>
<%end if%>
<%end select%>
</body>
</html>

