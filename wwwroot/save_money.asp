<!--#include file="session.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim checkflag,CompanyType
checkflag=false
if session("level")=10 or session("level")=7 then checkflag=True
CompanyType = GetFieldDataBySQL("select CompanyType from sysconfig","int",0)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<script src="Js/Calendar.js"></script>
<script language="javascript" src="inc/func.js" type="text/javascript"></script>
<script language="javascript">
function chk()
{
	if(!CheckIsNumericOrNull(document.form1.money,"����д��","�����д����")) return false;
	if(document.form1.type.value==""){
		alert("��ѡ���տ���Ŀ��");
		document.form1.type.focus();
		return false;
	}
	if(document.form1.type.value=="3" || document.form1.type.value=="4"){
		if(document.form1.userid.value==""){
			alert("��ѡ���տ�Ա����");
			document.form1.userid.focus();
			return false;
		}
	}
	if(!document.form1.wzsk.checked){
		symoney=parseFloat(document.getElementById("symoney"+document.form1.type.value).value);
		money=parseFloat(document.form1.money.value);
		if(symoney<money){
			alert("�տ������Χ.");
			return false;
		}
	}
	if(!CheckIsNull(document.form1.beizhu,"����д��ע˵����")) return false;
	if(!CheckIsNull(document.getElementById("times"),"����д�տ�ʱ�䣡")) return false;
	if(!CheckIsNull(document.form1.type,"��ѡ���տ����ͣ�")) return false;
	document.form1.btn_save.disabled=true;
}
function chktype(val)
{
	if(val=='3' || val=='4')
	{
		document.getElementById("userid").disabled = false;
	}
	else
	{
		document.getElementById("userid").disabled = true;
	}
}
</script>
<link href="Css/calendar-blue.css" rel="stylesheet">
<link href="admin/zxcss.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.STYLE2 {color: #999999}
.STYLE3 {color: #CCCCCC}
.STYLE4 {color: #FF0000}
-->
</style>
</head>

<body>
<%
if session("level")="" then
response.Write "<script>alert('�Բ���,��ûȨ�޽����ҳ�棡');history.go(-1)</script>"
response.end 
end if%>
<p>&nbsp;</p>
<% select case request("action")
case "added"
dim userid,peplename,typetext
select case request("type")
case 1
	typetext = "��ϵ�ɷ�"
	money11=request("money")-request("no_save1")
case 2
	typetext = "ѡƬ���ѽɷ�"
	money11=request("money")-request("no_save2")
case 3
	typetext = "�������ѽɷ�"
	money11=request("money")-request("no_save3")
case 4
	typetext = "������ѽɷ�"
	money11=request("money")-request("no_save4")
end select

if session("username")<>"" then
	userid=session("userid")
	peplename=session("username")
else
	userid=conn.execute("select userid from shejixiadan where id="&xmid)(0)
	peplename=conn.execute("select peplename from yuangong where username='"&userid&"'")(0)
end if
if money11>0 then
response.Write "<script>alert('�Բ����տ����Ӧ�ս�������տ������Ƿ�ѡ����ȷ��');history.go(-1)</script>"
Response.End
end if
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from save_money",conn,1,3
rs.addnew
if request("userid")<>"" then
	rs("userid")=request("userid")
else
	rs("userid")=userid
end if
rs("group")=conn.execute("select [group] from yuangong where username='"&userid&"'")(0)
rs("xiangmu_id")=request("id")
rs("money")=request("money")
rs("type")=request("type")
if request("wzsk")=1 then
	rs("wzsk")=1
end if
rs("times")=request("times")
rs("beizhu")=htmlencode2(request("beizhu"))
rs.update
temp = rs.bookmark
rs.bookmark = temp
save_id=rs("ID")
rs.close

rs.open "select * from hesuan where times=#"&datevalue(request("times"))&"# and [userid]='"&userid&"'",conn,1,3
if rs.eof then
rs.addnew
rs("times")=datevalue(request("times"))
rs("userid")=userid
rs.update
end if
rs.close
set rs=nothing

conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&request("id")&",0,'"&userid&"','"&peplename&" ��� "&typetext&" "&request("money")&"Ԫ','������',#"&now()&"#)")

Call FinalMoneySum(Cint(request("id")),True)
if request("btn_save")="�ύ��ת����ӡ" then
	if request("type")=1 then
		response.Redirect("paizhao_print.asp?id="&save_id&"&xiangmu_id="&request("id"))
	else
		response.Redirect("save_money_print.asp?id="&save_id)
	end if
else
	response.Redirect("save_money.asp?id="&request("id"))
end if
case "dksave"
If Request.Form("money")="" Then
Response.Write "<script>alert('�ֿڽ���Ϊ��!');history.go(-1)</script>"
Response.End
End If
if session("level")=1 then
userid=session("userid")
else
userid=conn.execute("select userid from shejixiadan where id="&request("id")&"")(0)
end if
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from save_dk",conn,1,3
rs.addnew
rs("userid")=userid
rs("group")=conn.execute("select [group] from yuangong where username='"&userid&"'")(0)
rs("xiangmu_id")=request("id")
rs("money")=request("money")
rs("times")=now()
rs("beizhu")=htmlencode2(request("beizhu"))
rs.update
rs.close
set rs=nothing
response.Redirect("save_money.asp?id="&request("id")&"")

case "dele"
	if session("username")<>"" then
		userid=session("userid")
		peplename=session("username")
	else
		userid=conn.execute("select userid from shejixiadan where id="&xmid)(0)
		peplename=conn.execute("select peplename from yuangong where username='"&userid&"'")(0)
	end if
	
	xm_id = Cint(request("id2"))
	
	dim rssave
	set rssave = conn.execute("select * from save_money where id="&request("id"))
	if not (rssave.eof and rssave.bof) then
		select case rssave("type")
		case 1
			typetext = "��ϵ�ɷ�"
		case 2
			typetext = "ѡƬ���ѽɷ�"
		case 3
			typetext = "�������ѽɷ�"
		case 4
			typetext = "������ѽɷ�"
		end select
		
		conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&xm_id&",0,'"&userid&"','"&peplename&" ɾ������ "&xm_id&" "&typetext&"��"&rssave("money")&"Ԫ��','������',#"&now()&"#)")
		conn.execute("delete from save_money where id="&request("id"))
		
		'conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&xm_id&",0,'"&userid&"','"&peplename&" ��ԤԼ���� "&xm_id&" "&typetext&"��"&rssave("money")&"Ԫ���������վ.','������',#"&now()&"#)")
		'conn.execute("update save_money set isdelete=true where id="&request("id"))
		
		Call FinalMoneySum(Cint(xm_id),True)
		response.Redirect "save_money.asp?id="&request("id2")
	else
		response.write "<script language='javascript'>alert('ɾ��ʧ��,δ�ҵ��˿���.');"
		response.write "location.hef='save_money.asp?id="&request("id2")&"';</script>"
		response.end()
	end if
	rssave.close
	set rssave = nothing
	
case "del"
if session("level")=3  then
response.Write "<script>alert('�Բ�����û��Ȩ�޽��иò���!');history.go(-1)</script>"
Response.End
end if
conn.execute("delete from save_dk where id="&request("id")&"")
response.Redirect "save_money.asp?id="&request("id2")&""
Response.End
%>
<%case "add"%>
<form action="save_money.asp?action=added" method="post" name="form1" onSubmit="return chk()">
<table width="96%" height="177"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#FFFFFF">
    <td height="18" colspan="2"><div align="left">���ţ�<%=request("id")%> ������������<%=conn.execute("select lxpeple from kehu where id=(select kehu_id from shejixiadan where id="&request("id")&")")(0)%><br>
      <span class="style1">
      <%
		set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where id="&request("id")&"",conn,1,1
		fujia_money=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)
		if isnull(fujia_money) then fujia_money=0
		%>
      </span>�ܽ��: <span class="style1">
      <%fujia2_money=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
		if isnull(fujia2_money) then fujia2_money=0
goumai_money=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
		if isnull(goumai_money) then goumai_money=0


if not isnull(conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)) then
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)
	else
	money1=0
	end if
	if not isnull(conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)) then
	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
	else
	money2=0
	end if
	if not isnull(conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)) then
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
	else
	money3=0
	end if
	money4=rs("jixiang_money")
	money=money1+money2+money3+money4
	response.Write money
	taoxi_save=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&" and type=1")(0)
	fujia_save=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&" and type=2")(0)
	fujia2_save=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&" and type=3")(0)
	goumai_save=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&" and type=4")(0)
	if isnull(taoxi_save) then taoxi_save=0
	if isnull(fujia_save) then fujia_save=0
	if isnull(fujia2_save) then fujia2_save=0
	if isnull(goumai_save) then goumai_save=0
	
	no_save1=money4-jixiang_save
	no_save2=money1-fujia1_save
	no_save3=money2-fujia2_save
	no_save4=money3-goumai_save
	%>
      </span> Ԫ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ܿ�δ�ɣ�<font color=red><%=(fujia_money-fujia_save)+(money4-taoxi_save)+(fujia2_money-fujia2_save)+(goumai_money-goumai_save)%></font>Ԫ&nbsp;&nbsp;<br>
��ϵ�<%=money4%>/δ��:<span class="style4"><%=money4-taoxi_save%>
<input name="symoney1" type="hidden" id="symoney1" value="<%=money4-taoxi_save%>">
Ԫ</span>&nbsp;&nbsp;&nbsp;ѡƬ����:<%=fujia_money%>/δ��:<span class="style4"><%=fujia_money-fujia_save%><input name="symoney2" type="hidden" id="symoney2" value="<%=fujia_money-fujia_save%>">
Ԫ&nbsp;</span>&nbsp;&nbsp;��������:<%=fujia2_money%>/δ��:<span class="style4"><%=fujia2_money-fujia2_save%>
<input name="symoney3" type="hidden" id="symoney3" value="<%=fujia2_money-fujia2_save%>">
Ԫ<%if CompanyType=0 then%>&nbsp;</span>&nbsp;&nbsp;�������:<%=goumai_money%>/δ��:<span class="style4"><%=goumai_money-goumai_save%>
<input name="symoney4" type="hidden" id="symoney4" value="<%=goumai_money-goumai_save%>">Ԫ</span><%end if%></div></td>
    </tr>
  <tr bgcolor="#FFFFFF">
    <td width="71%" height="39">&nbsp;</td>
    <td width="29%"><input name="no_save1" type="hidden" id="no_save1" value="<%=no_save1%>">
      <input name="no_save2" type="hidden" id="no_save2" value="<%=no_save2%>">
      <input name="no_save3" type="hidden" id="no_save3" value="<%=no_save3%>">
      <input name="no_save4" type="hidden" id="no_save4" value="<%=no_save4%>"></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="64" colspan="2"><div align="center">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="9%"><div align="right">�����</div></td>
        <td width="41%"><input name="money" type="text" id="money" size="6"  >          <select name="type" id="type" onChange="javascript:chktype(this.value)">
            <option value="">��ѡ��</option>
            <option value="1">��ϵ�ɷ�</option>
            <option value="2">ѡƬ���ѽɷ�</option>
            <option value="3">�������ѽɷ�</option>
            <%if CompanyType=0 then%><option value="4">������ѽɷ�</option><%end if%>
          </select>
           �տ
           <%Call ShowUserSelect("userid", "1,4,5,14", "username", "��ѡ��...", "", 0, true)%></td>
        <td width="24%"><%if CheckOldMoneyControl() then%>ʱ��
          <input name="times" type="text" id="times" size="15" value="<%=now%>" readonly>
          <span class="font"><a onClick="return showCalendar('times', 'y-mm-dd');" href="#"><img src="Image/Button.gif" id="IMG2" align="absMiddle" border="0" /></a></span>
          <%else
		  response.write "<input name='times' type='hidden' id='times' value='"&now()&"'>"
		end if%></td>
        <td width="26%">ˢ���տ    
          <input type="checkbox" name="wzsk" value="1" style="border:none"></td>
      </tr>
      <tr>
        <td valign="top"><div align="right">��ע˵����</div></td>
        <td colspan="3"><textarea name="beizhu" cols="100" rows="7" id="beizhu"></textarea></td>
      </tr>
    </table>
    </div></td>
  </tr>
</table>
  <br>
  <table width="83%" height="40"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="center">
  <input name="btn_save" type="submit" id="btn_save" value="���ύ">
  &nbsp; 
  <input name="btn_save" type="submit" id="btn_save" value="�ύ��ת����ӡ">
&nbsp;  
<input type="button" name="Submit" value="����" onClick="javascript:history.go(-1)">
  <input type="hidden" name="id" id="id" value="<%=request("id")%>">
      </div></td>
    </tr>
  </table>
</form>
  <br>
  <br>
  <table width="83%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="21%" valign="top"><div align="right">�����</div></td>
      <td width="79%"><textarea name="yongyu" cols="65" rows="7" id="yongyu"><%if not isnull(conn.execute("select yongyu from save_yongyu where userid='"&"admin"&"'")(0)) then 
	  response.Write encode2(conn.execute("select yongyu from save_yongyu where userid='"&"admin"&"'")(0))
	  end if%>
  </textarea></td>
    </tr>
  </table>
  <table width="600"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>      <%set rs=conn.execute("select distinct(dated) from (select datevalue(times) as dated from fujia where  xiangmu_id="&request("id")&" union  select datevalue(times) as dated from fujia2  where xiangmu_id="&request("id")&")")
if not rs.eof then response.Write "<strong>���ӹ���</strong><br>"
while not rs.eof
fujia1_money=0
fujia2_money=0
response.Write rs("dated")&":&nbsp;&nbsp;"
set rs2=conn.execute("select * from fujia where xiangmu_id="&request("id")&" and datevalue(times)=#"&rs("dated")&"#")
		if not rs2.eof then
		while not rs2.eof 
		response.Write conn.execute("select yunyong from yunyong where id="&rs2("jixiang")&"")(0)&" "
		response.write "������"&rs2("sl")&"&nbsp;&nbsp;"
		response.Write "��"&rs2("money")&"Ԫ&nbsp;&nbsp;��ע��"&encode(rs2("beizhu"))&"&nbsp;&nbsp;&nbsp;&nbsp;"
		fujia1_money=fujia1_money+rs2("money")
		rs2.movenext
		wend 
		end if
		rs2.close
		set rs2=nothing
		set rs2=conn.execute("select * from fujia2 where xiangmu_id="&request("id")&" and datevalue(times)=#"&rs("dated")&"#")
		if not rs2.eof then
		while not rs2.eof 
		response.Write conn.execute("select yunyong from yunyong where id="&rs2("jixiang")&"")(0)&" "
		response.Write "���:"&rs2("money")&"Ԫ&nbsp;&nbsp;��ע��"&encode(rs2("beizhu"))&"&nbsp;&nbsp;&nbsp;"
		fujia2_money=fujia2_money+rs2("money")
		rs2.movenext
		wend 
		end if
		response.Write "&nbsp;&nbsp;&nbsp;<strong>�ϼ��ܽ��:</strong><font color=red>"&fujia1_money+fujia2_money&"</font>Ԫ"
		rs2.close
		set rs2=nothing
		response.Write "<br>"
rs.movenext
wend
rs.close
set rs=nothing%></td>
    </tr>
  </table>
<% case "dk" %>
<form action="save_money.asp?action=dksave" method="post" name="form1">
<table width="96%" height="177"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#FFFFFF">
    <td height="18" colspan="2"><div align="left">���ţ�<%=request("id")%> ������������<%=conn.execute("select lxpeple from kehu where id=(select kehu_id from shejixiadan where id="&request("id")&")")(0)%><br>
      �ܽ��:
          <%if not isnull(conn.execute("select sum(money) from fujia where xiangmu_id="&request("id")&"")(0)) then
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&request("id")&"")(0)
	else
	money1=0
	end if
	if not isnull(conn.execute("select sum(money) from fujia2 where xiangmu_id="&request("id")&"")(0)) then
	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&request("id")&"")(0)
	else
	money2=0
	end if
	if not isnull(conn.execute("select sum(money) from goumai where xiangmu_id="&request("id")&"")(0)) then
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&request("id")&"")(0)
	else
	money3=0
	end if
	money4=conn.execute("select jixiang_money from shejixiadan where id="&request("id")&"")(0)


	if isnull(conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&"")(0)) then
	money5=0 
	else
	money5=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&"")(0)
	end if
	response.Write "<font color=red>"&money1+money2+money3+money4&"</font>"%>
          Ԫ&nbsp;&nbsp;&nbsp;<strong>�ܿ�δ�ɣ�<font color=red><%=money1+money2+money3+money4-money5%></font><span class="style1">Ԫ</span></strong><span class="style1">&nbsp; <br>
          ��ϵδ��:
          <%jixiang_save=conn.execute("select sum(money) from save_money where type=1 and xiangmu_id="&request("id")&"")(0)
		if isnull(jixiang_save) then jixiang_save=0
		no_save1=money4-jixiang_save
		response.Write "<font color=red>"&no_save1&"</font>"
		%>
          Ԫ&nbsp;&nbsp;ѡƬ����δ��:
          <%fujia1_save=conn.execute("select sum(money) from save_money where type=2 and xiangmu_id="&request("id")&"")(0)
		if isnull(fujia1_save) then fujia1_save=0
		no_save2=money1-fujia1_save
		response.Write "<font color=red>"&no_save2&"</font>"%>
          Ԫ&nbsp;��������δ��:
          <%fujia2_save=conn.execute("select sum(money) from save_money where type=3 and xiangmu_id="&request("id")&"")(0)
		if isnull(fujia2_save) then fujia2_save=0
		no_save3=money2-fujia2_save
		response.Write "<font color=red>"&no_save3&"</font>"%>
          Ԫ<%if CompanyType=0 then%>&nbsp;&nbsp;�������δ��:
          <%goumai_save=conn.execute("select sum(money) from save_money where type=4 and xiangmu_id="&request("id")&"")(0)
if isnull(goumai_save) then goumai_save=0
no_save4=money3-goumai_save
response.Write "<font color=red>"&no_save4&"</font>"
%>
Ԫ<%end if%></span></div></td>
    </tr>
  <tr bgcolor="#FFFFFF">
    <td width="71%" height="39"><div align="center"></div>
    <div align="left"></div></td>
    <td width="29%"><input name="no_save1" type="hidden" id="no_save1" value="<%=no_save1%>">
      <input name="no_save2" type="hidden" id="no_save2" value="<%=no_save2%>">
      <input name="no_save3" type="hidden" id="no_save3" value="<%=no_save3%>">
      <input name="no_save4" type="hidden" id="no_save4" value="<%=no_save4%>"></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="64" colspan="2"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="21%"><div align="right">�ֿ۽�</div></td>
        <td><input name="money" type="text" id="money" size="6"  > 
  (Ԫ)</td>
        </tr>
      <tr>
        <td valign="top"><div align="right">��ע˵����</div></td>
        <td><textarea name="beizhu" cols="65" rows="7" id="beizhu"></textarea></td>
      </tr>
    </table>
    </div></td>
  </tr>
</table>
  <table width="83%" height="40"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="center">
  <input type="submit" name="Submit" value="�ύ">
&nbsp;&nbsp;&nbsp;
  <input type="button" name="Submit" value="����" onClick="javascript:history.go(-1)">
  <input type="hidden" name="id" value="<%=request("id")%>">
      </div></td>
    </tr>
  </table>
</form>
  <table width="83%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="21%" valign="top"><div align="right">�����</div></td>
      <td width="79%"><textarea name="yongyu" cols="65" rows="7" id="yongyu"><%if not isnull(conn.execute("select yongyu from save_yongyu where userid='"&"admin"&"'")(0)) then 
	  response.Write encode2(conn.execute("select yongyu from save_yongyu where userid='"&"admin"&"'")(0))
	  end if%>
  </textarea></td>
    </tr>
  </table>
  <table width="600"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>      <%set rs=conn.execute("select distinct(dated) from (select datevalue(times) as dated from fujia where  xiangmu_id="&request("id")&" union  select datevalue(times) as dated from fujia2  where xiangmu_id="&request("id")&")")
if not rs.eof then response.Write "<strong>���ӹ���</strong><br>"
while not rs.eof
fujia1_money=0
fujia2_money=0
response.Write rs("dated")&":&nbsp;&nbsp;"
set rs2=conn.execute("select * from fujia where xiangmu_id="&request("id")&" and datevalue(times)=#"&rs("dated")&"#")
		if not rs2.eof then
		while not rs2.eof 
		response.Write conn.execute("select yunyong from yunyong where id="&rs2("jixiang")&"")(0)&" "
		response.Write "���:"&rs2("money")&"Ԫ&nbsp;&nbsp;��ע��"&encode(rs2("beizhu"))&"&nbsp;&nbsp;&nbsp;&nbsp;"
		fujia1_money=fujia1_money+rs2("money")
		rs2.movenext
		wend 
		end if
		rs2.close
		set rs2=nothing
		set rs2=conn.execute("select * from fujia2 where xiangmu_id="&request("id")&" and datevalue(times)=#"&rs("dated")&"#")
		if not rs2.eof then
		while not rs2.eof 
		response.Write conn.execute("select yunyong from yunyong where id="&rs2("jixiang")&"")(0)&" "
		response.Write "���:"&rs2("money")&"Ԫ&nbsp;&nbsp;��ע��"&encode(rs2("beizhu"))&"&nbsp;&nbsp;&nbsp;"
		fujia2_money=fujia2_money+rs2("money")
		rs2.movenext
		wend 
		end if
		response.Write "&nbsp;&nbsp;&nbsp;<strong>�ϼ��ܽ��:</strong><font color=red>"&fujia1_money+fujia2_money&"</font>Ԫ"
		rs2.close
		set rs2=nothing
		response.Write "<br>"
rs.movenext
wend
rs.close
set rs=nothing%></td>
    </tr>
  </table>
  <%case else
  conn.execute("update shejixiadan s inner join save_money m on s.id=m.xiangmu_id set m.userid=s.userid where isnull(m.userid) or m.userid=''")
  %>
  <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr>
    <td width="79%"><div align="left">&nbsp;&nbsp;<span class="style1">
      <%
		set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where id="&request("id")&"",conn,1,1
		fujia_money=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)
		if isnull(fujia_money) then fujia_money=0
		%>
    </span>�ܽ��:
        <span class="style1">
        <%fujia2_money=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
		if isnull(fujia2_money) then fujia2_money=0
goumai_money=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
		if isnull(goumai_money) then goumai_money=0


if not isnull(conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)) then
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)
	else
	money1=0
	end if
	if not isnull(conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)) then
	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
	else
	money2=0
	end if
	if not isnull(conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)) then
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
	else
	money3=0
	end if
	money4=rs("jixiang_money")
	if isnull(conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&"")(0)) then
	money5=0 
	else
	money5=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&"")(0)
	end if
	if isnull(conn.execute("select sum(money) from save_dk where xiangmu_id="&request("id")&"")(0)) then
	money6=0 
	else
	money6=conn.execute("select sum(money) from save_dk where xiangmu_id="&request("id")&"")(0)
	end if
	money=money1+money2+money3+money4
	response.Write money
	taoxi_save=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&" and type=1")(0)
	fujia_save=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&" and type=2")(0)
	fujia2_save=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&" and type=3")(0)
	goumai_save=conn.execute("select sum(money) from save_money where xiangmu_id="&request("id")&" and type=4")(0)
	if isnull(taoxi_save) then taoxi_save=0
	if isnull(fujia_save) then fujia_save=0
	if isnull(fujia2_save) then fujia2_save=0
	if isnull(goumai_save) then goumai_save=0
	%>
        </span>
        
        Ԫ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ܿ�δ�ɣ�<font color=red><%=(fujia_money-fujia_save)+(money4-taoxi_save)+(fujia2_money-fujia2_save)+(goumai_money-goumai_save)%></font><span class="style1">Ԫ&nbsp;&nbsp;&nbsp;<span class="STYLE2">&nbsp;&nbsp;ԤԼ�㶨�� �ڵ��ӡ�� &nbsp;</span><br>
&nbsp;&nbsp;��ϵ�<%=money4%>/δ��:<span class="style4"><%=money4-taoxi_save%>Ԫ</span>&nbsp;&nbsp;&nbsp;ѡƬ����:<%=fujia_money%>/δ��:<span class="style4"><%=fujia_money-fujia_save%>Ԫ&nbsp;</span>&nbsp;&nbsp;��������:<%=fujia2_money%>/δ��:<span class="style4"><%=fujia2_money-fujia2_save%>Ԫ<%if CompanyType=0 then%>&nbsp;</span>&nbsp;&nbsp;�������:<%=goumai_money%>/δ��:<span class="style4"><%=goumai_money-goumai_save%>Ԫ</span><%end if%></span></div></td>
    <td width="9%"><a href="save_money.asp?action=dk&id=<%=request("id")%>">�ֿ�ƾ֤</a></td>
    <td width="12%"><strong><a href="save_money.asp?action=add&id=<%=request("id")%>"><span class="STYLE4">����տ�</span></a>&nbsp;</strong></td>
  </tr>
</table>
  <br>
  <br>

<%
dim rsdj,djid
set rsdj = conn.execute("select id from save_money where xiangmu_id="&request("id")&" and [type]=1 order by times")
if not (rsdj.eof and rsdj.bof) then
	djid = rsdj("id")
else
	djid = 0
end if
rsdj.close
set rsdj = nothing
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from save_money where xiangmu_id="&request("id")&" and isdelete=false order by times desc",conn,1,1
if not rs.eof then
%>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <%
  i=0
  while not rs.eof 
  i=i+1%>
  <tr bgcolor="#efefef">
    <td width="56%">&nbsp;&nbsp;��<font color="#FF0000"><%=rs("money")%></font>&nbsp;(Ԫ)[
      <%if rs("type")=1 then
	response.Write "��ϵ�ɷ�"
	elseif rs("type")=2 then
	response.Write "ѡƬ���ѽɷ�"
	elseif rs("type")=3 then
	response.Write "�������ѽɷ�"
	elseif rs("type")=4 then
	response.Write "������ѽɷ�"
	end if%>
      ]
      <%if rs("wzsk")=1 then response.write "<font color=blue>[ˢ���տ�]</font>"%>      &nbsp;&nbsp;<%=rs("times")%></td><td width="44%" align="right">&nbsp;&nbsp;&nbsp;
        <%
	  if rs("type")<>1 then
	  	response.write "<a href='save_money_print.asp?id="&rs("id")&"'>"
		if rs("ischeck")=0 and checkflag then response.write "<span id='sp_"&rs("id")&"'>���</span>"
		response.write "��ױ/ѡƬ</a>&nbsp;&nbsp;"
	  else
	  	'response.write "<a href='save_money_print.asp?id="&rs("id")&"'></a>"
	  	if rs("id")=djid then
			response.write "<a href='dinjin_print.asp?id="&rs("id")&"&dj=1&xiangmu_id="&request("id")&"'>"
			if rs("ischeck")=0 and checkflag then response.write "<span id='sp_"&rs("id")&"'>���</span>"
			response.write "����</a>&nbsp;&nbsp;"
		else 
			response.write "<a href='paizhao_print.asp?id="&rs("id")&"&xiangmu_id="&request("id")&"'><span class='style1'>"
			if rs("ischeck")=0 and checkflag then response.write "<span id='sp_"&rs("id")&"'>���</span>"
			response.write "���ս��ӡ</span></a>&nbsp;&nbsp;&nbsp;"
		end if
		response.write "<a href='dinjin_print.asp?Dj=2&id="&rs("id")&"&xiangmu_id="&request("id")&"'><span class='STYLE2'>��ϸ����</span><span class='STYLE3'>��</span></a>&nbsp;&nbsp;"
      end if
      if session("level")=10 or instr(session("level2"),"721")>0 then
	  	response.write "&nbsp;&nbsp;<a href='save_money.asp?action=dele&id="&rs("id")&"&id2="&request("id")&"' onClick=""return confirm('ȷ��Ҫɾ���˿�����')"">ɾ��</a>&nbsp;&nbsp;"
		'ȷ��Ҫ�����������������վ��
	  end if%>
	</td>
  </tr>

  <tr bgcolor="#FFFFFF">
    <td colspan="2"><span  style="line-height:20px">&nbsp;&nbsp;��ע��<%=replace(rs("beizhu"),"����","����")%><br>  &nbsp;&nbsp;�����ˣ�
  <%
	set rs_k = conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")
	if not rs_k.eof then
		response.write rs_k("peplename")
	else
		response.write "N/A"
	end if
	rs_k.close()
	set rs_k = nothing
	%> 
 &nbsp;&nbsp;
 <%
	if rs("orderid")>0 then
		select case rs("type")
			case 2
				sql = "select * from fujia where id="&rs("orderid")
			case 3
				sql = "select * from fujia2 where id="&rs("orderid")
			case 4
				sql = "select * from goumai where id="&rs("orderid")
				
		end select
		set rsjs = server.createobject("adodb.recordset")
		rsjs.open sql,conn,1,1
		if not (rsjs.eof and rsjs.bof) then
			response.write "�����ˣ�"
			if rsjs("userid")<>"" and not isnull(rsjs("userid")) then
				response.write conn.execute("select peplename from yuangong where username='"&rsjs("userid")&"'")(0)
			else
				response.write "��"
			end if
			if rsjs("userid2")<>"" and not isnull(rsjs("userid2")) then
				response.write "/"&conn.execute("select peplename from yuangong where username='"&rsjs("userid2")&"'")(0)
			end if
			if rs("type")=3 or rs("type")=4 then
				if rsjs("userid3")<>"" and not isnull(rsjs("userid3")) then
					response.write "/"&conn.execute("select peplename from yuangong where username='"&rsjs("userid3")&"'")(0)
				end if
			end if
		end if
		rsjs.close
		set rsjs = nothing
	end if
	response.write "&nbsp;&nbsp;<span id='sp_checkuser"&rs("id")&"'>"
	if rs("checkuserid")>0 then
		response.write "�����:"&GetFieldDataBySQL("select peplename from yuangong where id="&rs("checkuserid"),"str","N/A")
	end if
	response.write "</span>"
	%>
    </span></td>
  </tr>
  <%rs.movenext
  wend 
  end if
  rs.close
  set rs=nothing%>
</table>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="right">�����ܽ�<font color="#FF0000"><%=money5%></font>&nbsp;(Ԫ)&nbsp;&nbsp;</div></td>
  </tr>
</table>

<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from save_dk where xiangmu_id="&request("id")&" order by times asc",conn,1,1
if not rs.eof then
%>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <%while not rs.eof 
  i=i+1%>
  <tr bgcolor="#efefef">
    <td width="35%"><div align="left">&nbsp;&nbsp;��<font color="#FF0000"><%=rs("money")%></font>&nbsp;(Ԫ)[�ֿ۽��]</div></td>
    <td width="24%"><div align="left">&nbsp;&nbsp;<%=rs("times")%></div></td>
    <td width="41%"><div align="center"><a href="save_dk_print.asp?id=<%=rs("id")%>">��ӡ�ֿ�ƾ֤</a>&nbsp;
        <%if session("level")=10 or instr(session("level2"),"721")>0 then%>
	<a href="save_money.asp?action=del&id=<%=rs("id")%>&id2=<%=request("id")%>" onClick="return confirm('ȷ��Ҫɾ����')">ɾ��	</a>
	<%end if%>
	</div></td>
  </tr>

  <tr bgcolor="#FFFFFF">
    <td colspan="3"><span  style="line-height:20px">&nbsp;&nbsp;��ע��<%=rs("beizhu")%><br>
  &nbsp;&nbsp;�����ˣ�
  <%
	set rs_k = conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")
	if not rs_k.eof then
		response.write rs_k("peplename")
	else
		response.write "N/A"
	end if
	rs_k.close()
	set rs_k = nothing
	%>
    </span></td>
  </tr>
  <%rs.movenext
  wend 
  end if
  rs.close
  set rs=nothing%>
</table>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="right">�ֿ��ܽ�<font color="#FF0000"><%=money6%></font>&nbsp;(Ԫ)&nbsp;&nbsp;</div></td>
  </tr>
</table>
<%end select%>
<br>

<table width="80%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="80%"><%set rs=conn.execute("select distinct(dated) from (select datevalue(times) as dated from fujia where  xiangmu_id="&request("id")&" union  select datevalue(times) as dated from fujia2  where xiangmu_id="&request("id")&")")
i=0
while not rs.eof
i=i+1
response.Write "<a href='admin/xiadan_print2.asp?times="&rs("dated")&"&px="&i&"&id="&request("id")&"'>��"&i&"�μ�����ӡ</a>&nbsp;&nbsp;"
'<a href='jiamai_print.asp?times="&rs("dated")&"&px="&i&"&id="&request("id")&"'>С����ӡ</a>["&day(rs("dated"))&"��]&nbsp;&nbsp;&nbsp;" 
rs.movenext
wend
rs.close
set rs=nothing%></td>
    <td width="20%">&nbsp;</td>
  </tr>
</table>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bordercolor="#66FF33" bgcolor="#66FF33">
    <td width="66%">&nbsp;&nbsp;<span class="STYLE2">������ݺ��̨������������������˹�̨�����տ�</span></td>
    <td width="18%" height="24"><a href="admin/two_yongyu.asp">����������</a></td>
    <td width="16%"><div align="center"><a href="fujia.asp?action=add&id=<%=request("id")%>">���ѡƬ����</a></div></td>
  </tr>
</table>
<div align="center">
  <%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from fujia where xiangmu_id="&request("id")&"",conn,1,1
if rs.eof then
response.Write "<font color=red size=2>û������ ������ѡƬ����</font>"
else
%>
</div>
<table width="96%" height="41"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <%while not rs.eof%>
  <tr bgcolor="#FFFFFF">
    <td width="21%" height="20">&nbsp;&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rs("jixiang")&"")(0)%></td>
    <td width="16%">&nbsp;Ա����
        <%
	response.write conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)
	if rs("userid2")<>"" and not isnull(rs("userid2")) then
		response.write "/"&conn.execute("select peplename from yuangong where username='"&rs("userid2")&"'")(0)
	end if
	%></td>
    <td width="16%">&nbsp;������<%response.write rs("sl")
	if rs("pagevol")>0 then response.write "&nbsp;&nbsp;P����"& rs("pagevol")
	%></td>
    <td width="12%">&nbsp;���ã�<%=rs("money")%>(Ԫ)</td>
    <td width="35%"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="77%" height="20"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td>&nbsp;<%=rs("times")%></td>
            <td>&nbsp;
                  <% if rs("userid")<> conn.execute("select userid from shejixiadan where id="&request("id")&"")(0) and session("level")=1 then
				response.Write "["&conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)&"]"
				end if%>
            </td>
          </tr>
        </table></td>
      <td width="23%"><%if session("level")=10 or instr(session("level2"),"721")>0 then%>
              <a href="fujia.asp?action=dele1&id2=<%=rs("id")%>&id=<%=request("id")%>" onClick="return confirm('ȷ��Ҫɾ����')">ɾ��</a>
              <%end if%>
        </td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="18" colspan="5"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="7%" valign="top"><div align="right">��ע:</div></td>
        <td width="93%"><%=encode(rs("beizhu"))%> </td>
      </tr>
    </table></td>
  </tr>
  <%rs.movenext 
  wend
  rs.close
  set rs=nothing%>
</table>
<p>
  <%end if%>
</p>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#efefef">
  <tr>
    <td width="74%" height="24" bgcolor="#66FF33"><div align="right"><a href="admin/save3_yongyu.asp">��ױ������</a></div></td>
    <td width="26%" bgcolor="#66FF33"><div align="center"><a href="fujia.asp?action=add2&id=<%=request("id")%>">�����������</a></div></td>
  </tr>
</table>
<div align="center">
  <%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from fujia2 where xiangmu_id="&request("id")&"",conn,1,1
if rs.eof then
response.Write "<font color=red size=2>��û���������Ѽ�¼</font>"
else
%>
</div>
<table width="96%" height="43"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <%while not rs.eof%>
  <tr bgcolor="#FFFFFF">
    <td height="20">&nbsp;&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rs("jixiang")&"")(0)%></td>
    <td>&nbsp;Ա����
    <%
	if rs("userid")<>"" and not isnull(rs("userid")) then
		response.write conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)
	else
		response.write "��"
	end if
	if rs("userid2")<>"" and not isnull(rs("userid2")) then response.write "/"&conn.execute("select peplename from yuangong where username='"&rs("userid2")&"'")(0)
	if rs("userid3")<>"" and not isnull(rs("userid3")) then response.write "/"&conn.execute("select peplename from yuangong where username='"&rs("userid3")&"'")(0)
	%></td>
    <td>&nbsp;���ã�<%=rs("money")%>Ԫ</td>
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="52%">&nbsp;<%=rs("times")%></td>
        <td width="27%">����:<%=rs("sl")%></td>
    <td width="21%"><%if session("level")=10 or instr(session("level2"),"721")>0 then%>
              <a href="fujia.asp?action=dele2&id2=<%=rs("id")%>&id=<%=request("id")%>" onClick="return confirm('ȷ��Ҫɾ����')">ɾ��</a>
              <%end if%></td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="18" colspan="4"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="7%" valign="top"><div align="right">��ע:</div></td>
        <td width="93%"><%=encode(rs("beizhu"))%> </td>
      </tr>
    </table></td>
  </tr>
  <%rs.movenext 
  wend
  rs.close
  set rs=nothing%>
</table>
<br>
<%end if
  if CompanyType=0 then%>
<br>
<br>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#efefef">
  <tr bgcolor="#66FF33">
    <td width="74%" height="24"><div align="right"><a href="hz_beizhu.asp?id=<%=request("id")%>">�޸Ľ�黯ױʱ����Ҫ��</a></div></td>
    <td width="26%" bgcolor="#66FF33"><div align="center"><a href="fujia.asp?action=add3&id=<%=request("id")%>">��ӽ��ױ����</a></div></td>
  </tr>
</table>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;&nbsp;��黯ױ����ʱ��:
        <%response.Write conn.execute("select hz_time from shejixiadan where id="&request("id")&"")(0)&"&nbsp;&nbsp;"
	shijian=conn.execute("select hz from shejixiadan where id="&request("id")&"")(0)
	response.Write shijian
	 
	%></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;��ױҪ��:<%=conn.execute("select hz_beizhu from shejixiadan where id="&request("id")&"")(0)%></td>
  </tr>
</table>

<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from goumai where xiangmu_id="&request("id")&"",conn,1,1
if rs.eof then
response.Write "<font color=red size=2>��û�н�黯ױ���Ѽ�¼</font>"
else
%>
<table width="96%" height="43"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <%while not rs.eof%>
  <tr bgcolor="#FFFFFF">
    <td height="20">&nbsp;&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rs("jixiang")&"")(0)%></td>
    <td>&nbsp;Ա����
      <%
	if rs("userid")<>"" and not isnull(rs("userid")) then
		response.write conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)
	else
		response.write "��"
	end if
	if rs("userid2")<>"" and not isnull(rs("userid2")) then response.write "/"&conn.execute("select peplename from yuangong where username='"&rs("userid2")&"'")(0)
	if rs("userid3")<>"" and not isnull(rs("userid3")) then response.write "/"&conn.execute("select peplename from yuangong where username='"&rs("userid3")&"'")(0)
	%></td>
    <td>&nbsp;���ã�<%=rs("money")%>Ԫ</td>
    <td>&nbsp;������<%=rs("sl")%></td>
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="63%">&nbsp;<%=rs("times")%></td>
        <td width="37%"><%if session("level")=10 or instr(session("level2"),"721")>0 then%>
              <a href="fujia.asp?action=dele3&id2=<%=rs("id")%>&id=<%=request("id")%>" onClick="return confirm('ȷ��Ҫɾ����')">ɾ��</a>
          <%end if%></td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="18" colspan="5"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="7%" valign="top"><div align="right">��ע:</div></td>
        <td width="93%"><%=encode(rs("beizhu"))%> </td>
      </tr>
    </table></td>
  </tr>
  <%rs.movenext 
  wend
  rs.close
  set rs=nothing%>
</table>
<br>
<%end if
end if
if dg_invis then%>
<br>
<br>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#efefef">
  <tr bgcolor="#FFFFFF">
    <td width="60%" height="18" bgcolor="#66FF33"><div align="right">
      <%if session("level")=10 then%>
      <a href="sell.asp?action=type&id2=<%=request("id")%>"><strong>������</strong>
        <%end if%>
      </a> <a href="admin/two_yongyu.asp">&nbsp;&nbsp;&nbsp;</a> </div></td>
    <td width="20%" bgcolor="#66FF33" style="display:none"><div align="center"><a href="sell.asp?action=add&id=<%=request("id")%>">�����Ӱʦ������¼</a></div></td>
    <td width="20%" bgcolor="#66FF33"><div align="center"><a href="sells.asp?action=add&id=<%=request("id")%>">�����������¼</a></div></td>
  </tr>
</table>
<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sell_jilu where xiangmu_id="&request("id")&" order by times desc",conn,1,1
if rs.eof then%>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center">��û������</div></td>
  </tr>
</table>
<%else%>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
  <%
do while not rs.eof 
if rs("yuangong_id")<>"" then
Set yuangong_rs=conn.execute("select peplename from yuangong where Id="&rs("yuangong_id")&"")
peplename=yuangong_rs("peplename")
yuangong_rs.close
set yuangong_rs=nothing
end if
%>
  <tr bgcolor="#FFFFFF">
    <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><font color="#FF0000">&nbsp;����Ա��<%=peplename%></font></td>
        <td align="right"><%if session("level")=10 then response.write "<a href=?action=deldg&id2="&rs("id")&"&id="&request("id")&" onClick=""return confirm('ȷ��Ҫɾ����')"">ɾ��</a>"%>
          &nbsp; </td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>&nbsp;��Ʒ:<%=rs("name")%></td>
    <td width="15%">&nbsp;����:<%=rs("dj")%>Ԫ</td>
    <td width="15%">&nbsp;&nbsp;����:<%=rs("sl")%></td>
    <td width="25%">&nbsp;ʱ��:<%=rs("times")%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="4">&nbsp;��ע:<%=encode(rs("beizhu"))%></td>
  </tr>
  <%
  rs.movenext
  i=i+1
  if i>=count then exit do
  loop
  %>
</table>
<%end if
end if
%>
</body>
</html>