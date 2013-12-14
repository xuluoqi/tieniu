<!--#include file="connstr.asp"-->
<!--#include file="session.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
<!--
.df {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
}
-->
</style>
<script language="javascript" src="inc/func.js" type="text/javascript"></script>
<link href="zxcss.css" rel="stylesheet" type="text/css">
<link href="admin/zxcss.css" rel="stylesheet" type="text/css">
<script src="Js/Calendar.js"></script>
<link href="Css/TestDate.css" rel="stylesheet">
<link href="Css/calendar-blue.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<script language="javascript">
function chk()
{
if(!CheckIsNull(document.form1.jixiang,"请选择套系！")) return false;
if(!CheckIsNull(document.form1.money,"请填写套系金额！")) return false;
if(!CheckIsNull(document.form1.beizhu,"请填写下单具体说明！"))return false;
}
</script>　
<body>
<%select case request("action")
case "edited"
if request("hz_time")<>"" and request("hz")="" then
  response.Write "<script>alert('请选择化妆具体时间！');history.go(-1)</script>"
  Response.End
  end if
  
sys=conn.execute("select [CpMaxNum] from sysconfig")(0)
if isnull(sys) then sys=0
if request("pz_time")<>"" then
	sy_number=conn.execute("select count(*) from shejixiadan where (pz_time=#"&request("pz_time")&"# or pz_time2=#"&request("pz_time")&"#) and id<>"&request("id"))(0)
	if sy_number>=sys and sys<>0 then
		response.Write "<script> alert('摄影日期["& request("pz_time") &"]已达到最高摄影人数,请另选择摄影日期！');history.go(-1) </script>"
		response.end  
	end if
end if
if request("pz_time2")<>"" then
	sy_number=conn.execute("select count(*) from shejixiadan where (pz_time=#"&request("pz_time2")&"# or pz_time2=#"&request("pz_time2")&"#) and id<>"&request("id"))(0)
	if sy_number>=sys and sys<>0 then
		response.Write "<script> alert('摄影日期["& request("pz_time2") &"]已达到最高摄影人数,请另选择摄影日期！');history.go(-1) </script>"
		response.end  
	end if
end if
if request("kj_time")<>"" then
	kynum=conn.execute("select kyMaxNum from sysconfig")(0)
	if isnull(kynum) then kynum=0
	ky_number=conn.execute("select count(*) from shejixiadan where kj_time=#"&request("kj_time")&"#")(0)
	if ky_number>=kynum and kynum<>0 then
		response.Write "<script> alert('选片当天已达到最高选片人数,请另选择选片日期！');history.go(-1) </script>"
		response.end  
	end if
end if

if request("pz_time")<>"" and request("pz")="" then
  response.Write "<script>alert('请选择拍照具体时间！');history.go(-1)</script>"
  Response.End
  end if
  if request("pz_time2")<>"" and request("pz2")="" then
  response.Write "<script>alert('请选择拍照2具体时间！');history.go(-1)</script>"
  Response.End
  end if
  if request("hhz_time")<>"" and request("hhz")="" then
  response.Write "<script>alert('请选择回婚妆具体时间！');history.go(-1)</script>"
  Response.End
  end if
   if request("qj_time")<>"" and request("qj")="" then
  response.Write "<script>alert('请选择取件具体时间！');history.go(-1)</script>"
  Response.End
  end if
   if request("kj_time")<>"" and request("kj")="" then
  response.Write "<script>alert('请选择选片具体时间！');history.go(-1)</script>"
  Response.End
  end if
    if request("xg_time")<>"" and request("xg")="" then
  response.Write "<script>alert('请选择看版具体时间！');history.go(-1)</script>"
  Response.End
  end if
   if not isnumeric(trim(request("money"))) then
  response.Write "<script>alert('请填写金额！');history.go(-1)</script>"
  Response.End
  end if
  if request("beizhu")="" then
  	if request("beizhu2")="" then
  	  beizhu11="&nbsp;"
    else
  	  beizhu11=request("beizhu2")
    end if
  else
  	beizhu11=request("beizhu")
  end if
   
  if trim(request("danhao"))<>"" then
    danhao=conn.execute("select count(*) from shejixiadan where danhao='"&request("danhao")&"' and danhao<>'' and id<>"&request("id")&"")(0)
  	if danhao>0 then
   		response.Write "<script>alert('该单号已经存在，请检查单号是否错误！');history.go(-1)</script>"
  		Response.End
  	end if
  end if
  set rs2=server.CreateObject("adodb.recordset")
  rs2.open "select * from shejixiadan where id="&request("id"),conn,1,3
  rs2("danhao")=request("danhao")
  rs2("yx_cp_name")=request("yx_cp_name")
  rs2("yx_cp_name2")=request("yx_cp_name2")
  rs2("yx_cp_name3")=request("yx_cp_name3")
  rs2("yx_cp_memo")=request("yx_cp_memo")
  rs2("yx_hz_name")=request("yx_hz_name")
  rs2("yx_xg_name")=request("yx_xg_name")
  rs2("yx_ky_name")=request("yx_ky_name")
  rs2("yx_jhz_name")=request("yx_jhz_name")
  rs2("yx_jhlf_name")=request("yx_jhlf_name")
rs2("beizhu")=htmlencode2(beizhu11)
if request("hz_time")<>"" and request("hz")<>"" then
	rs2("hz_time")=request("hz_time")
	rs2("hz")=request("hz")
else
	rs2("hz_time")=null
	rs2("hz")=null
end if
if request("pz_time")<>"" and request("pz")<>"" then
	rs2("pz_time")=request("pz_time")
	rs2("pz")=request("pz")
else
	rs2("pz_time")=null
	rs2("pz")=null
end if 
if request("pz_time2")<>"" and request("pz2")<>"" then
	rs2("pz_time2")=request("pz_time2")
	rs2("pz2")=request("pz2")
else
	rs2("pz_time2")=null
	rs2("pz2")=null
end if 
if request("hhz_time")<>"" and request("hhz")<>"" then
	rs2("hhz_time")=request("hhz_time")
	rs2("hhz")=request("hhz")
else
	rs2("hhz_time")=null
	rs2("hhz")=null
end if 
if request("pzlf_time")<>"" and request("pzlf")<>"" then
	rs2("pzlf_time")=request("pzlf_time")
	rs2("pzlf")=request("pzlf")
else
	rs2("pzlf_time")=null
	rs2("pzlf")=null
end if 
if request("jhlf_time")<>"" and request("jhlf")<>"" then
	rs2("jhlf_time")=request("jhlf_time")
	rs2("jhlf")=request("jhlf")
else
	rs2("jhlf_time")=null
	rs2("jhlf")=null
end if 
if request("kj_time")<>"" and request("kj")<>"" then
	rs2("kj_time")=request("kj_time")
	rs2("kj")=request("kj")
else
	rs2("kj_time")=null
	rs2("kj")=null
end if
if request("qj_time")<>"" and request("qj")<>"" then
	rs2("qj_time")=request("qj_time")
	rs2("qj")=request("qj")
else
	rs2("qj_time")=null
	rs2("qj")=null
end if
if request("xg_time")<>"" and request("xg")<>"" then
	rs2("xg_time")=request("xg_time")
	rs2("xg")=request("xg")
else
	rs2("xg_time")=null
	rs2("xg")=null
end if

if request("o_hz_time")<>"" then
	if request("hz_time")<>"" and request("hz")<>"" then
		if cdate(request("hz_time"))<>cdate(request("o_hz_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","hz_time|hz",request("o_hz_time")&"|"&request("o_hz"),request("hz_time")&"|"&request("hz"))
			Call EditedTimeSaveToReport(request("id"),e,"hz",request("o_hz_time"),request("hz_time"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","hz_time|hz",request("o_hz_time")&"|"&request("o_hz"),request("hz_time")&"|"&request("hz"))
		Call EditedTimeSaveToReport(request("id"),e,"hz",request("o_hz_time"),request("hz_time"))
	end if
end if
if request("o_pz_time")<>"" then
	if request("pz_time")<>"" and request("pz")<>"" then
		if cdate(request("pz_time"))<>cdate(request("o_pz_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","pz_time|pz",request("o_pz_time")&"|"&request("o_pz"),request("pz_time")&"|"&request("pz"))
			Call EditedTimeSaveToReport(request("id"),e,"pz",request("o_pz_time"),request("pz_time"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","pz_time|pz",request("o_pz_time")&"|"&request("o_pz"),request("pz_time")&"|"&request("pz"))
		Call EditedTimeSaveToReport(request("id"),e,"pz",request("o_pz_time"),request("pz_time"))
	end if
end if
if request("o_pzlf_time")<>"" then
	if request("pzlf_time")<>"" and request("pzlf")<>"" then
		if cdate(request("pzlf_time"))<>cdate(request("o_pzlf_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","pzlf_time|pzlf",request("o_pzlf_time")&"|"&request("o_pzlf"),request("pzlf_time")&"|"&request("pzlf"))
			Call EditedTimeSaveToReport(request("id"),e,"pzlf",request("o_pzlf_time"),request("pzlf_time"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","pzlf_time|pzlf",request("o_pzlf_time")&"|"&request("o_pzlf"),request("pzlf_time")&"|"&request("pzlf"))
		Call EditedTimeSaveToReport(request("id"),e,"pzlf",request("o_pzlf_time"),request("pzlf_time"))
	end if
end if
if request("o_jhlf_time")<>"" then
	if request("jhlf_time")<>"" and request("jhlf")<>"" then
		if cdate(request("jhlf_time"))<>cdate(request("o_jhlf_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","jhlf_time|jhlf",request("o_jhlf_time")&"|"&request("o_jhlf"),request("jhlf_time")&"|"&request("jhlf"))
			Call EditedTimeSaveToReport(request("id"),e,"jhlf",request("o_jhlf_time"),request("jhlf_time"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","jhlf_time|jhlf",request("o_jhlf_time")&"|"&request("o_jhlf"),request("jhlf_time")&"|"&request("jhlf"))
		Call EditedTimeSaveToReport(request("id"),e,"jhlf",request("o_jhlf_time"),request("jhlf_time"))
	end if
end if
if request("o_kj_time")<>"" then
	if request("kj_time")<>"" and request("kj")<>"" then
		if cdate(request("kj_time"))<>cdate(request("o_kj_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","kj_time|kj",request("o_kj_time")&"|"&request("o_kj"),request("kj_time")&"|"&request("kj"))
			Call EditedTimeSaveToReport(request("id"),e,"kj",request("o_kj_time"),request("kj_time"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","kj_time|kj",request("o_kj_time")&"|"&request("o_kj"),request("kj_time")&"|"&request("kj"))
		Call EditedTimeSaveToReport(request("id"),e,"kj",request("o_kj_time"),request("kj_time"))
	end if
end if
if request("o_qj_time")<>"" then
	if request("qj_time")<>"" and request("qj")<>"" then
		if cdate(request("qj_time"))<>cdate(request("o_qj_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","qj_time|qj",request("o_qj_time")&"|"&request("o_qj"),request("qj_time")&"|"&request("qj"))
			Call EditedTimeSaveToReport(request("id"),e,"qj",request("o_qj_time"),request("qj_time"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","qj_time|qj",request("o_qj_time")&"|"&request("o_qj"),request("qj_time")&"|"&request("qj"))
		Call EditedTimeSaveToReport(request("id"),e,"qj",request("o_qj_time"),request("qj_time"))
	end if
end if
if request("o_xg_time")<>"" then
	if request("xg_time")<>"" and request("xg")<>"" then
		if cdate(request("xg_time"))<>cdate(request("o_xg_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","xg_time|xg",request("o_xg_time")&"|"&request("o_xg"),request("xg_time")&"|"&request("xg"))
			Call EditedTimeSaveToReport(request("id"),e,"xg",request("o_xg_time"),request("xg_time"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","xg_time|xg",request("o_xg_time")&"|"&request("o_xg"),request("xg_time")&"|"&request("xg"))
		Call EditedTimeSaveToReport(request("id"),e,"xg",request("o_xg_time"),request("xg_time"))
	end if
end if

rs2("jixiang_money")=request("money")
rs2("stated")=request("stated")
if request("sl2")<>"" and isnumeric(request("sl2")) then rs2("sl2")=request("sl2")

rs2.update
rs2.close
set rs2=nothing
response.Write "<script>alert('修改下单成功!');location='xiadan.asp?id="&request("id")&"&action=edit'</script>"

case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where id="&request("id"),conn,1,1
%><form action="xiadan.asp?action=edited&id=<%=request("id")%>" method="post"  name="form1">
<table width="97%" height="281" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="xu_kuan">
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" align="right" valign="middle" class="font">套系：</td>
      <td valign="middle" class="font"><%=conn.execute("select jixiang from jixiang where id="&rs("jixiang")&"")(0)%>
      <input name="jixiang" type="hidden" id="jixiang" value="<%=rs("jixiang")%>"></td>
      <td align="right" class="font">套系金额：</td>
      <td valign="middle" class="font"><input name="money" type="text" id="money" size="13" readonly value="<%=rs("jixiang_money")%>">
（元）&nbsp;&nbsp;&nbsp;&nbsp;原价：<%=conn.execute("select old_money from jixiang where id="&rs("jixiang")&"")(0)%></td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="31" align="right" valign="middle" class="font">*摄影日期1：</td>
      <td height="31" colspan="3" valign="middle" class="font"><input name="pz_time" type="text" maxlength="10" id="txtAwardDate" size="13" value="<%=rs("pz_time")%>">
        <a onClick="return showCalendar('pz_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
      <input name="pz" type="text" size="3" value="<%=rs("pz")%>">&nbsp;点
      <select name="yx_cp_name" id="yx_cp_name">
        <option value="">预设摄影师1</option>
        <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where [level]=4 and isdisabled=0",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_cp_name") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
      </select>
      <select name="yx_cp_name2" id="yx_cp_name2">
        <option value="">预设摄影师2</option>
        <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where [level]=4 and isdisabled=0",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_cp_name2") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
      </select>
      <select name="yx_cp_name3" id="yx_cp_name3">
        <option value="">预设摄影师3</option>
        <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where [level]=4 and isdisabled=0",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_cp_name3") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
      </select>
      <input name="o_pz_time" type="hidden" id="o_pz_time" value="<%=rs("pz_time")%>">
<input name="o_pz" type="hidden" id="o_pz" value="<%=rs("pz")%>"></td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td align="right" class="font">摄影日期2：</td>
      <td height="30" colspan="3" class="font"><input name="pz_time2" type="text" maxlength="10" id="pz_time2" size="13" value="<%=rs("pz_time2")%>">
        <a onClick="return showCalendar('pz_time2', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="pz2" type="text" size="3" value="<%=rs("pz2")%>">
      &nbsp;点&nbsp;&nbsp;&nbsp;摄影/化妆备注
<input name="yx_cp_memo" type="text" id="yx_cp_memo" value="<%=rs("yx_cp_memo")%>" size="30"></td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td align="right" class="font">拍照礼服：</td>
      <td class="font"><input name="pzlf_time" type="text" maxlength="10" id="pzlf_time" size="13" value="<%=rs("pzlf_time")%>" />
          <a onClick="return showCalendar('pzlf_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
          <input name="pzlf" type="text" size="3" value="<%=rs("pzlf")%>">
  点
  <select name="yx_hz_name" id="yx_hz_name">
    <option value="">预设拍照妆</option>
    <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where [level]=5 and isdisabled=0",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_hz_name") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
  </select>
  <input name="o_pzlf_time" type="hidden" id="o_pzlf_time" value="<%=rs("pzlf_time")%>">
  <input name="o_pzlf" type="hidden" id="o_pzlf" value="<%=rs("pzlf")%>"></td>
      <td height="30" align="right" class="font">*结婚礼服：</td>
      <td height="30" class="font"><input name="jhlf_time" type="text" id="jhlf_time" size="13" value="<%=rs("jhlf_time")%>" >
          <a onClick="return showCalendar('kj_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="jhlf" type="text" id="jhlf" value="<%=rs("jhlf")%>" size="3">
        点
        <select name="yx_jhlf_name" id="yx_jhlf_name">
          <option value="">预设礼服师</option>
          <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where ([level]=11 or [level]=14) and isdisabled=0 order by [level],username",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_jhlf_name") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
        </select>
        <input name="o_jhlf_time" type="hidden" id="o_jhlf_time" value="<%=rs("jhlf_time")%>">
        <input name="o_jhlf" type="hidden" id="o_jhlf" value="<%=rs("jhlf")%>"></td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
	  <td height="30" align="right" valign="middle" class="font">*选片日期：</td>
	  <td height="30" valign="middle" class="font"><input name="kj_time" type="text" id="kj_time" size="13" value="<%=rs("kj_time")%>" >
        <a onClick="return showCalendar('kj_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="kj" type="text" size="3" value="<%=rs("kj")%>">
点
<select name="yx_ky_name" id="yx_ky_name">
  <option value="">预设选片门市</option>
  <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where [level]=1 and isdisabled=0",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_ky_name") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
</select>
<input name="o_kj_time" type="hidden" id="o_kj_time" value="<%=rs("kj_time")%>">
<input name="o_kj" type="hidden" id="o_kj" value="<%=rs("kj")%>"></td>
      <td height="30" align="right" class="font">*结婚化妆：</td>
      <td height="30" class="font"><input name="hz_time" type="text" id="hz_time" size="13"  value="<%=rs("hz_time")%>">
          <a onClick="return showCalendar('hz_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="hz" type="text" size="3" value="<%=rs("hz")%>">
        点
        <select name="yx_jhz_name" id="yx_jhz_name">
          <option value="">预设结婚妆</option>
          <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where ([level]=5 or [level]=14) and isdisabled=0 order by [level]",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_jhz_name") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
        </select>
        <input name="o_hz_time" type="hidden" id="o_hz_time" value="<%=rs("hz_time")%>">
        <input name="o_hz" type="hidden" id="o_hz" value="<%=rs("hz")%>"></td>
    </tr>
	<tr align="left" valign="middle" bgcolor="#ffffff">
	  <td height="30" align="right" valign="middle" class="font">*看版日期：</td>
	  <td height="30" valign="middle" class="font"><input name="xg_time" type="text" id="xg_time" size="13"  value="<%=rs("xg_time")%>">
        <a onClick="return showCalendar('xg_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="xg" type="text" size="3" value="<%=rs("xg")%>">
        点
        <select name="yx_xg_name" id="yx_xg_name">
          <option value="">预设看版人员</option>
          <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where ([level]=1 or [level]=2) and isdisabled=0 order by [level]",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_xg_name") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
        </select>
        <input name="o_xg_time" type="hidden" id="o_xg_time" value="<%=rs("xg_time")%>">
        <input name="o_xg" type="hidden" id="o_xg" value="<%=rs("xg")%>"></td>
	  <td height="30" align="right" class="font">回婚妆：</td>
      <td height="30" class="font"><input name="hhz_time" type="text" id="hhz_time" size="13"  value="<%=rs("hhz_time")%>">
        <a onClick="return showCalendar('hhz_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="hhz" type="text" size="3" value="<%=rs("hhz")%>">
点</td>
	</tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" align="right" class="font">取件日期：</td>
      <td height="30" colspan="3" class="font"><input name="qj_time" type="text" id="qj_time" size="13"  value="<%=rs("qj_time")%>">
        <a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="qj" type="text" size="3" value="<%=rs("qj")%>">
点
<input name="o_qj_time" type="hidden" id="o_qj_time" value="<%=rs("qj_time")%>">
<input name="o_qj" type="hidden" id="o_qj" value="<%=rs("qj")%>"></td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" colspan="4" class="font">&nbsp;&nbsp;手动单号:
        <input name="danhao" type="text" id="danhao" size="8" value="<%=rs("danhao")%>">        &nbsp;&nbsp;毛片回件情况:
        <input name="stated" type="radio" value="1"  <%if rs("stated")=1 then response.Write "checked"%>>
正常
<input type="radio" name="stated" value="2" <%if rs("stated")=2 then response.Write "checked"%>>
急
<input type="radio" name="stated" value="3" <%if rs("stated")=3 then response.Write "checked"%>>
特急&nbsp;&nbsp;&nbsp;&nbsp;拍摄多款选
<input name="sl2" type="text" id="sl2" size="7" value="<%=rs("sl2")%>">
张</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="14" colspan="4" class="font"><div align="left">&nbsp;&nbsp;套系应有:
          <%
	  id=split(rs("yunyong"),", ")
	  sl=split(rs("sl"),", ")
	  for ii=lbound(id) to ubound(id)
	 	 response.write "<a href=javascript:javascript:openViewPic('admin/yunyong_pic.asp?action=view&id="&id(ii)&"',800,600) title='点击查看套系图片'>"&conn.execute("select yunyong from yunyong where id="&id(ii)&"")(0)&"</a>["&sl(ii)&"]&nbsp;&nbsp;&nbsp;"
		 if ii=6 or ii=12 then
		 response.Write"<br>&nbsp;&nbsp;"
		 end if
	  next
	  
	  %>
</div></td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="15" colspan="4" class="font">&nbsp;</td>
    </tr>
</table>
<br>
  <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="22%" valign="top"><div align="right">备注说明： </div></td>
      <td width="78%"><textarea name="beizhu" cols="70" rows="7" id="beizhu" <% if session("level")=1 or session("level")=7 then response.Write("disabled") End If %>><%=encode2(rs("beizhu"))%></textarea><% if session("level")=1 or session("level")=7 then response.write "<input name='beizhu2' type='hidden' id='beizhu2' value='"&encode2(rs("beizhu"))&"'>"%></td>
    </tr>
  </table>
  <div align="center">  <table width="97%" height="47"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="center">
      
	  <input name="tijiao" type="submit" id="确定" value="确定" onClick="return chk();">
  　　　　	
    <input name="reset" type="button" id="reset" value="返回" onClick="javascript:history.go(-1)">
    <input name="id" type="hidden" id="id" value="<%=request("id")%>">
    <input name="yunyong" type="hidden" id="yunyong" value="<%=rs("yunyong")%>">
    <input name="shuliang" type="hidden" id="shuliang" value="<%=rs("sl")%>">
</div></td>
    </tr>
  </table>
  </div>
</form>
<%rs.close
set rs=nothing%>
<%case else%> 

<form action="xiadan_save.asp" method="post"  name="form1" onSubmit="return check()">
<table width="97%" height="271" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="xu_kuan">

    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" class="font"><div align="right">另选择下单时间：</div></td>
      <td class="font"><input name="times" type="text" id="times" value="<%=date%>" size="13" readonly>
        <a onClick="return showCalendar('times', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a></td>
      <td colspan="2" class="font">&nbsp;&nbsp;&nbsp;如果没另选择下单时间，默认时间为添加当天日期</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td width="124" height="30" class="font"><div align="right">摄影类型：</div></td>
      <td class="font">        
	  <select name="type" id="type" onChange="changelocation(document.form1.type.options[document.form1.type.selectedIndex].value)">
	  <option value=""><%if request("ids")<>"" then 
	 response.Write conn.execute("select companytype from companytype where id="&conn.execute("select type from jixiang where id="&request("ids")&"")(0)&"")(0)
	 else
	 response.Write "请选择"
	 end if%></option>
	  <%set rs=conn.execute("select * from companytype")
	  while not rs.eof%>
	  <option value="<%=rs("id")%>"><%=rs("companytype")%></option>
	  <%rs.movenext
	  wend
	  rs.close
	  set rs=nothing%>
      </select>
套系：
<select name="jixiang" id="jixiang" size="1" ONCHANGE="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}">
 <option value="<%if request("ids")<>"" then 
 response.Write request("ids")
 end if %>"><%if request("ids")<>"" then 
 response.Write conn.execute("select jixiang from jixiang where id="&request("ids"))(0)
 else
 response.Write "请选择套系"
 end if %></option>
 
  <%set rs1=server.CreateObject("adodb.recordset")
	rs1.open "select * from jixiang where ishidden=0 order by [type],px",conn,1,1
	while not rs1.eof%>
  <option value="xiadan.asp?ids=<%=rs1("id")%>&id=<%=request("id")%>"><%=rs1("jixiang")%></option>
  <%rs1.movenext 
		wend 
		rs1.close
		set rs1=nothing%>
</select></td>
      <td width="116" class="font"><div align="right">套系金额：</div></td>
      <td class="font"><input name="money" type="text" id="money" size="10" value="<%if request("ids")<>"" then 
	  response.Write conn.execute("select money from jixiang where id="&request("ids")&"")(0)
	  end if%>" onKeyUp="value=value.replace(/[^\d]/g,'')   "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
      （元）&nbsp;&nbsp;&nbsp;预收金额：
      <input name="savemoney" type="text" id="savemoney" value="0" size="6">
元</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="31" align="right" class="font">摄影日期1：</td>
      <td width="233" height="31" class="font"><input name="pz_time" type="text" maxlength="10" id="txtAwardDate" size="13"/ >
        <a onClick="return showCalendar('pz_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
      <input name="pz" type="text" size="3">
&nbsp;点</td>
      <td width="116" align="right" class="font">拍照礼服：</td>
      <td width="279" class="font"><input name="pzlf_time" type="text" maxlength="10" id="pzlf_time" size="13" />
          <a onClick="return showCalendar('pzlf_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
          <input name="pzlf" type="text" size="3">
  &nbsp;点 (可为空)</td>
    </tr>
	<tr>
	  <td height="30" align="right" bgcolor="#FFFFFF" class="font">摄影日期2：</td>
	  <td height="30" class="font" bgcolor="#FFFFFF"><input name="pz_time2" type="text" maxlength="10" id="pz_time2" size="13"/ >
        <a onClick="return showCalendar('pz_time2', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="pz2" type="text" size="3">
&nbsp;点</td>
	  <td height="30" align="right" valign="middle" bgcolor="#FFFFFF"class="font">结婚礼服：</td>
	  <td height="30" align="left" valign="middle" bgcolor="#FFFFFF" class="font"><input name="jhlf_time" type="text" id="jhlf_time" size="13" />
          <a onClick="return showCalendar('jhlf_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" name="IMG2" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
          <input name="jhlf" type="text" id="jhlf" size="3">
  &nbsp;点 (可为空)</td>
	</tr>
	<tr> 
	<td height="30" class="font" bgcolor="#FFFFFF"><div align="right">看版日期：</div></td>
      <td height="30" class="font" bgcolor="#FFFFFF">
        <input name="xg_time" type="text" id="xg_time23" size="13" >
      <a onClick="return showCalendar('xg_time', 'y-mm-dd');" href="#"> <img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
      <input name="xg" type="text" size="3">
&nbsp;点</td>
	  <td height="30" align="right" valign="middle" bgcolor="#ffffff" class="font">结婚化妆：</td>
	  <td height="30" align="left" valign="middle" bgcolor="#ffffff" class="font"><a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#">
        <input name="hz_time" type="text" maxlength="10" id="hz_time" size="13" >
        </a><a onClick="return showCalendar('hz_time', 'y-mm-dd');" href="#"> <img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="hz" type="text" size="3">
  &nbsp;点</td>
	</tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
	<td height="30" class="font"><div align="right">选片日期：</div></td>
      <td height="30" class="font"><input name="kj_time" type="text" id="sjyq7" size="13" >
        <a onClick="return showCalendar('kj_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="kj" type="text" size="3" value="">
&nbsp;点</td>
      <td height="30" align="right" class="font">回婚妆：</td>
      <td height="30" class="font"><a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#">
        <input name="hhz_time" type="text" maxlength="10" id="hhz_time" size="13" >
        </a><a onClick="return showCalendar('hhz_time', 'y-mm-dd');" href="#"> <img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="hhz" type="text" size="3">
  &nbsp;点</td>
    </tr>
	
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" align="right" bgcolor="#FFFFFF" class="font">预设选片门市：</td>
      <td height="30" bgcolor="#FFFFFF" class="font"><select name="yx_ky_name" id="yx_ky_name">
          <option value="">请选择...</option>
          <%
			  set rss = server.CreateObject("adodb.recordset")
			  rss.open "select * from yuangong where [level]=1 and isdisabled=0",conn,1,1
			  do while not rss.eof
			  %>
          <option value="<%=rss("peplename")%>"><%=rss("peplename")%></option>
          <%
			  rss.movenext
			  loop
			  rss.close
			  %>
      </select></td>
      <td height="30" align="right" class="font">取件日期：</td>
      <td height="30" class="font"><input name="qj_time" type="text" id="qj_time2" size="13">
          <a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="qj" type="text" size="3">
  &nbsp;点</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" colspan="4" class="font">&nbsp;&nbsp;预约单号:
        <input name="danhao" type="text" id="danhao" size="8">        
        &nbsp; 毛片回件情况:
        <input name="stated" type="radio" value="1" checked>
正常
<input type="radio" name="stated" value="2">
急
<input type="radio" name="stated" value="3">
特急&nbsp;&nbsp;&nbsp;拍摄多款选
<input name="sl2" type="text" id="sl2" size="7" value="<%if request("ids")<>"" then 
response.Write conn.execute("select sl2 from jixiang where id="&request("ids")&"")(0)
end if%>">
张</td>
    </tr>
	<%if request("ids")<>"" then%>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="20" colspan="4" class="font"><div align="left">&nbsp;&nbsp;套系应有:
          <%
		  id1=conn.execute("select yunyong from jixiang where id="&request("ids")&"")(0)
		  sl1=conn.execute("select sl from jixiang where id="&request("ids")&"")(0)
	  id=split(id1,", ")
	  sl=split(sl1,", ")
	 for ii=lbound(id) to ubound(id)
	 response.write "<a href=javascript:javascript:openViewPic('admin/yunyong_pic.asp?action=view&id="&id(ii)&"',800,600) title='点击查看套系图片'>"&conn.execute("select yunyong from yunyong where id="&id(ii)&"")(0)&"</a>["&sl(ii)&"]&nbsp;&nbsp;&nbsp;"
	 if ii=6 or ii=12 then
	 response.Write"<br>&nbsp;&nbsp;"
	 end if
	 next
	  
	  %>
</div></td>
    </tr>
	<%end if%>
</table>

  <%set rs2=server.CreateObject("adodb.recordset")
rs2.open "select * from yunyong_type where ishidden=0 order by px asc",conn,1,1%>
  <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
    <%while not rs2.eof 
  zz=zz
  %>
    <tr>
      <td bgcolor="#efefef">&nbsp;&nbsp;<strong><%=rs2("name")%></strong></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">
        <div align="center">
          <%set rs3=server.CreateObject("adodb.recordset")
	rs3.open "select * from yunyong where type_id="&rs2("id")&" and ishidden=0 order by px asc",conn,1,1
	if not rs3.eof then%>
        </div>
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
          <%
								do while not rs3.eof 
								
							%>
          <tr>
            <%
								for a=1 to 3
								zz=zz+1
								if not rs3.eof then
								i=i-1
								%>
            <td  align=center valign="top">
              <table width="226" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#666666">
                <tr>
                  <td width="174" align="center" bgcolor="#FFFFFF">
                    <div align="left">
                      <input name="check" type="checkbox" id="check" value="<%=rs3("id")%>" 
		  <%if request("ids")<>"" then
		  if instr(conn.execute("select yunyong from jixiang where id="&request("ids")&"")(0),rs3("id"))>0 then response.Write "checked"
		  end if
		%>>
                      <%
		if len(zz)=1 then
		response.Write "<font color=red>00"&zz&"</font>-"
		elseif len(zz)=2 then
		response.Write "<font color=red>0"&zz&"</font>-"	
		else
		response.Write "<font color=red>"&zz&"</font>-"	
		end if
		if rs3("pic")<>"" and not isnull(rs3("pic")) then
			response.write "<a href=javascript:javascript:openViewPic('admin/yunyong_pic.asp?action=view&id="&rs3("id")&"',800,600) title='点击查看套系图片'><font color=red>"&rs3("yunyong")&"</font></a>"
		else
			response.write rs3("yunyong")
		end if
		if rs3("type3")=1 then response.write "&nbsp;<font color=#999999>[礼服系列]</font>"
		%> </div></td>
                  <td width="52" align="center" bgcolor="#FFFFFF"><div align="left">
                      <input name="sl<%=rs3("id")%>" type="text" id="sl<%=rs3("id")%>" size="3" value="<%if request("ids")<>"" then
		 set rs4=conn.execute("select * from jixiang where id="&request("ids")&"") 
		  if instr(rs4("yunyong"),rs3("id"))>0 then 
		  tt=split(rs4("yunyong"),", ")
		  for y=lbound(tt) to ubound(tt)
		  if trim(tt(y))=trim(rs3("id")) then t3=y
		  next 
		  x=split(rs4("sl"),", ")
		  response.Write x(t3)
		  end if
		  
		  rs4.close
		  set rs4=nothing
		   end if%>">
                  </div></td>
                </tr>
            </table></td>
            <%
								  rs3.movenext
								  elseif rs3.eof then
								  response.write"<td height=20 width=250 align=center></td>"
								  end if
								  next
								  %>
          </tr>
          <tr> </tr>
          <%
								loop
								rs3.close
								set rs3=nothing
								%>
        </table>
        <div align="center">
          <%else
	response.Write "该类别暂时无数据"
	end if%>
      </div></td>
    </tr>
    <%rs2.movenext
  wend
  rs2.close
  set rs2=nothing%>
  </table>
  <br>
  <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="22%" valign="top"><div align="right">备注说明： </div></td>
      <td width="78%"><textarea name="beizhu" cols="70" rows="7" id="beizhu"><%if request("ids")<>"" then 
	  response.Write encode2(conn.execute("select beizhu from jixiang where id="&request("ids")&"")(0))
	  end if%></textarea></td>
    </tr>
  </table>
  <div align="center">  <table width="97%" height="51"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="center">
      <input name="queding" type="submit" id="确定" value="确定" onClick="return chk();">
  　　　　　　　　
    <input name="reset" type="button" id="reset" value="返回" onClick="javascript:history.go(-1)">
    <input name="id" type="hidden" id="id" value="<%=request("id")%>">
      </div></td>
    </tr>
  </table>
  </div>
</form>
<%end select%>
<%
dim rssearch
dim sqlsearch
set rssearch=server.createobject("adodb.recordset")
sqlsearch = "select * from jixiang where ishidden=0 order by [type],px"
rssearch.open sqlsearch,connstr,1,1
response.write"<script language = ""JavaScript"">"
response.write"var onecount;"
response.write"onecount=0;"
response.write"subcat = new Array();"
        count = 0
        do while not rssearch.eof 
	response.write"subcat["&count&"] = new Array('"& trim(rssearch("type"))&"','"& trim(rssearch("jixiang"))&"','xiadan.asp?ids="&trim(rssearch("id"))&"&id="&request("id")&"');"
	dim count
        count = count + 1
        rssearch.movenext
        loop
        rssearch.close
response.write"onecount="&count&";"
response.write"function changelocation(locationid)"
response.write"{"
response.write"document.form1.jixiang.length = 0;" 
response.write"var locationid=locationid;"
response.write"var i;"
response.write"document.form1.jixiang.options[0] = new Option('请选择套系','');"
response.write"for (i=0;i < onecount; i++)"
response.write"{"
response.write"if (subcat[i][0] == locationid)"
response.write"{"
response.write"document.form1.jixiang.options[document.form1.jixiang.length] = new Option(subcat[i][1], subcat[i][2]);"
response.write"}"
response.write"}"
response.write"}"
response.write"</script>"
dim rssearch2
dim sqlsearch2
set rssearch2=server.createobject("adodb.recordset")
sqlsearch2 = "select * from jixiang where ishidden=0 order by [type],px"
rssearch2.open sqlsearch2,connstr,1,1
response.write"<script language = ""JavaScript"">"
response.write"var onecount2;"
response.write"onecount2=0;"
response.write"subcat2 = new Array();"
        count2 = 0
        do while not rssearch2.eof 
	response.write"subcat2["&count2&"] = new Array('"& trim(rssearch2("id"))&"','"&trim(rssearch2("money"))&"','"&rssearch2("yunyong")&"');"
	dim count2
        count2 = count2 + 1
        rssearch2.movenext
        loop
        rssearch2.close
response.write"onecount2="&count2&";"
		set rssearch2=nothing
response.write"function changelocation2(locationid)"
response.write"{"
response.write"var locationid=locationid;"
response.write"var i;"
response.write"var kk='';"
response.write"var jj='';"
response.write"for (i=0;i < onecount2; i++)"
response.write"{"
response.write"if (subcat2[i][0] == locationid)"
response.write"{"
response.write"kk=subcat2[i][1];"
response.write"jj=subcat2[i][2];"
response.write"}"
response.write"}"
response.write"}"
response.write"</script>"
%>
<p>&nbsp;</p>
</body>
</html>

