<!--#include file="connstr.asp"-->
<!--#include file="session.asp"-->
<!--#include file="../inc/function.asp"-->
<%Response.Buffer=True%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<script src="../Js/Calendar.js"></script>
<script language="javascript" src="../inc/func.js" type="text/javascript"></script>
<script language="javascript">   
function chkSameDataRow(idlist){
	var tr_row;
	var arr = idlist.split(",");
	for(var k=0;k<arr.length;k++){
		if(arr[k]!=""){
			tr_row = document.all("tr_"+arr[k]);
			if(tr_row.length>1){
				for(var i=0;i<tr_row.length;i++){
					tr_row[i].style.fontWeight ="bold";
					tr_row[i].style.backgroundColor ="#EEEEEE";
				}
			}
		}
	}
}
function loadingHidden()
{
	eval("document.getElementById(\"loadingimg\").style.display=\"none\"");
}
function loadingShow()
{
	eval("document.getElementById(\"loadingimg\").style.display=\"\"");
}
</script>
<%
response.write "<script language='JavaScript'>"&vbcrlf
response.write "<!-- Begin"&vbcrlf
response.write "top.window.moveTo(-4,-4);"&vbcrlf
response.write "if (document.all) {"&vbcrlf
response.write "top.window.resizeTo(screen.availWidth+8,screen.availHeight+8);"&vbcrlf
response.write "}"&vbcrlf
response.write "else if (document.layers||document.getElementById) {"&vbcrlf
response.write "if (top.window.outerHeight<screen.availHeight||top.window.outerWidth&vbcrlf<screen.availWidth){"&vbcrlf
response.write "top.window.outerHeight = screen.availHeight;"&vbcrlf
response.write "top.window.outerWidth = screen.availWidth;"&vbcrlf
response.write "}"&vbcrlf
response.write "}"&vbcrlf
response.write "//  End -->"&vbcrlf
response.write "</script> "&vbcrlf
%>
<link href="../Css/calendar-blue.css" rel="stylesheet">
<link href="zxcss.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.jh {	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-style: normal;
	line-height: normal;
}
.style5 {color: #000000}
.style6 {font-size: 10px}
.STYLE7 {
	font-size: 16px;
	font-weight: bold;
}
.STYLE9 {color: #666666}
.STYLE10 {color: #999999}
.STYLE11 {color: #CCCCCC}
-->
</style>
</head>
<body>


<%
sub init_key()
	daogou_choucheng=0
	pz_choucheng11=0
	hz_choucheng11=0
	fujia_save11=0
	jixiang_choucheng=0
	jixiang_money=0
	jx_mymoney=0
	money414=0
	fujia_save11=0
	moeny113=0
	daogou_choucheng=0
	fujia_choucheng=0
	money_all=0
	sl2 = 0
	alldgmoney=0
	allhqmoney=0
	allpersonhq=0
	ReceivablesMoney=0
	RecFujiaMoney=0
	AllRecFujiaMoney=0
	hq_notsavemoney=0
	hq_allmoney=0
	hq_mymoney=0
	hq_minemoney=0
	hq_indate_savemoney=0
	hq_indate_allsavemoney=0
	all_tx_wed=0
	idlist=""
	msidlist=","
	money00=0
	money11=0
	money22=0
	money33=0
	money44=0
	
	MonthWedsuitCost=0
	AllWedsuitCost=0
	AllXiangmuMoney=0
	AllQiankuanMoney=0
	MonthFujiaCost=0
	AllFujiaCost=0
	AllFujiaMoney=0
	
	all_cpVolume=0
	all_txVolume=0
	
	fujia_hepai=0
	fujia_fenpai1=0
	fujia_fenpai2=0

	hqsave_hepai1=0
	hqsave_hepai2=0
	
	dd_all_dingjin=0
	dd_all_paizhao=0
end sub

dim userid,peplename
userid = request("userid")
set rspn = conn.execute("select peplename from yuangong where username='"&userid&"'")
if not rspn.eof then
	peplename = rspn("peplename")
end if
rspn.close
set rspn = nothing

yeard=request.form("year")
monthd=request.form("month")

fromtime = request.form("fromtime")
totime = request.form("totime")

dim datearea,sql_time
if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
	datearea = yeard&" �� "&monthd&" ��"
	sql_time = "not isnull(times) and datevalue(times)<#"&yeard&"-"&monthd&"-1# and not isnull(times)"
end if
if (fromtime<>"" and not isnull(fromtime)) and (totime<>"" and not isnull(totime)) then
	datearea = fromtime&" �� "&totime
	sql_time = "not isnull(times) and datevalue(times)<#"&fromtime&"# and not isnull(times)"
end if

qj_flag = request.form("qj_flag")
'if qj_flag="" then qj_flag="hidden"

function GetSqlCheckDateString(fieldname)
	if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
		GetSqlCheckDateString = " not isnull("&fieldname&") and year("&fieldname&")="&yeard&" And month("&fieldname&")="&monthd&" and not isnull("&fieldname&")"
	end if
	if (fromtime<>"" and not isnull(fromtime)) and (totime<>"" and not isnull(totime)) then
		GetSqlCheckDateString = " not isnull("&fieldname&") and datevalue("&fieldname&")>=#"&datevalue(fromtime)&"# And datevalue("&fieldname&")<=#"&datevalue(totime)&"# And not isnull("&fieldname&")"
	end if
end function

function GetNonSaveMoney(orderid,types)
	Dim rstmp
	'��ϵ��
	Dim z_jixiangmoney
	Set rstmp = conn.execute("select jixiang_money from shejixiadan where id="&orderid)
	If Not (rstmp.eof And rstmp.bof) Then
		z_jixiangmoney = rstmp(0)
	Else
		GetNonSaveMoney = 0 
		Exit Function 
	End If 
	
	Dim sqldate
	if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) Then
		dim t_dates
		t_dates = dateadd("d",-1,cdate(yeard&"-"&monthd+1&"-1"))
		sqldate = "datevalue(times)<=#"&t_dates&"#"
	End If 
	if (fromtime<>"" and not isnull(fromtime)) and (totime<>"" and not isnull(totime)) Then
		sqldate = "datevalue(times)<=#"&totime&"#"
	End If 
	
	Dim z_fujia, z_fujia2, z_goumai
	Dim z_jixiangsave, z_fujiasave, z_fujia2save, z_goumaisave

	'===============================================================

	If types = 0 Or types = 1 Then
		'��ǰʱ���ֹ��ϵ�ɿ�
		z_jixiangsave=conn.execute("select sum(money) from save_money where xiangmu_id="&orderid&" and [type]=1 and not isnull(times) and "&sqldate&" and not isnull(times)")(0)
		If IsNull(z_jixiangsave) Then z_jixiangsave = 0

		If types = 1 Then 
			GetNonSaveMoney = z_jixiangmoney - z_jixiangsave
			Exit Function 
		End If 
	End If 

	'===============================================================
	
	If types = 0 Or types = 2 Then
		'��ǰʱ���ֹ��������
		z_fujia=conn.execute("select sum(money) from fujia where xiangmu_id="&orderid&" and not isnull(times) and "&sqldate&" and not isnull(times)")(0)
		If IsNull(z_fujia) Then z_fujia = 0

		'��ǰʱ���ֹ���ڽɿ�
		z_fujiasave=conn.execute("select sum(money) from save_money where xiangmu_id="&orderid&" and [type]=2 and not isnull(times) and "&sqldate&" and not isnull(times)")(0)
		If IsNull(z_fujiasave) Then z_fujiasave = 0
		
		If types = 2 Then 
			GetNonSaveMoney = z_fujia - z_fujiasave
			Exit Function 
		End If 
	End If 

	'===============================================================
	
	If types = 0 Or types = 3 Then
		'��ǰʱ���ֹ��������
		z_fujia2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&orderid&" and not isnull(times) and "&sqldate&" and not isnull(times)")(0)
		If IsNull(z_fujia2) Then z_fujia2 = 0

		'��ǰʱ���ֹ����ױ�ɿ�
		z_fujia2save=conn.execute("select sum(money) from save_money where xiangmu_id="&orderid&" and [type]=3 and not isnull(times) and "&sqldate&" and not isnull(times)")(0)
		If IsNull(z_fujia2save) Then z_fujia2save = 0

		If types = 3 Then 
			GetNonSaveMoney = z_fujia2 - z_fujia2save
			Exit Function 
		End If 
	End If 

	'===============================================================
		
	If types = 0 Or types = 4 Then
		'��ǰʱ���ֹ��������
		z_goumai=conn.execute("select sum(money) from goumai where xiangmu_id="&orderid&" and not isnull(times) and "&sqldate&" and not isnull(times)")(0)
		If IsNull(z_goumai) Then z_goumai = 0

		'��ǰʱ���ֹ����ױ�ɿ�
		z_goumaisave=conn.execute("select sum(money) from save_money where xiangmu_id="&orderid&" and [type]=4 and not isnull(times) and "&sqldate&" and not isnull(times)")(0)
		If IsNull(z_goumaisave) Then z_goumaisave = 0

		If types = 4 Then 
			GetNonSaveMoney = z_goumai - z_goumaisave
			Exit Function 
		End If 
	End If 
	
	'��ǰʱ���ֹ��Ƿ��
	If types = 0 Then
		GetNonSaveMoney = (z_jixiangmoney + z_fujia + z_fujia2 + z_goumai) - (z_jixiangsave + z_fujiasave + z_fujia2save + z_goumaisave)
		Exit Function 
	End If 

end Function

Dim arr_cons_info()
dim arr_cons_minmoney(),arr_cons_maxmoney(),arr_cons_vol(),arr_cons_txsl()
dim rscons,cons_count,losttype_count,losttypecount,conscount
Dim rslosttype

Function InitConsInfo()
  Set rslosttype = server.CreateObject("adodb.recordset")
  rslosttype.open "select * from customerlosttype order by px asc",conn,1,1
  losttypecount = rslosttype.recordcount
  ReDim arr_cons_info(losttypecount, 5)
  losttype_count=0
  Do While Not rslosttype.eof
	  losttype_count = losttype_count + 1
	  set rscons = server.createobject("adodb.recordset")
	  rscons.open "select * from CustomerConsumption where typeid=1 and customerlosttypeid="&rslosttype("id")&" order by px asc",conn,1,1
	  conscount = rscons.recordcount
	  arr_cons_info(losttype_count,0) = rslosttype("id")
	  arr_cons_info(losttype_count,1) = rslosttype("title")

	  if not (rscons.eof and rscons.bof) then
		cons_count=0
		redim arr_cons_minmoney(conscount)
		redim arr_cons_maxmoney(conscount)
		redim arr_cons_vol(conscount)
		redim arr_cons_txsl(conscount)
		do while not rscons.eof
			cons_count = cons_count + 1
			arr_cons_minmoney(cons_count) = rscons("minmoney")
			arr_cons_maxmoney(cons_count) = rscons("maxmoney")
			arr_cons_vol(cons_count) = 0
			arr_cons_txsl(cons_count) = 0
			rscons.movenext
		Loop
		arr_cons_info(losttype_count,2) = arr_cons_minmoney
		arr_cons_info(losttype_count,3) = arr_cons_maxmoney
		arr_cons_info(losttype_count,4) = arr_cons_vol
		arr_cons_info(losttype_count,5) = arr_cons_txsl
	  else
		arr_cons_info(losttype_count,2) = null
		arr_cons_info(losttype_count,3) = null
		arr_cons_info(losttype_count,4) = Null
		arr_cons_info(losttype_count,5) = null
	  end if
	  rscons.close
	  set rscons = Nothing
	  rslosttype.movenext
  Loop 
  rslosttype.close
  Set rslosttype=Nothing 
End Function 

if (yeard="" or isnull(yeard)) and (monthd="" or isnull(monthd)) and (fromtime="" or isnull(fromtime)) and (totime="" or isnull(totime)) then
	response.write "<div style='width:100%; text-align:center'><br><br><br><br><br>����ѡ��ʱ��Σ��ٽ��в�ѯ��</div>"

else
	set rsyg = server.CreateObject("adodb.recordset")
	rsyg.open "select * from yuangong where username='"&userid&"'",conn,1,1
	if not rsyg.eof then
		typed=rsyg("level")
		cur_peplename=rsyg("peplename")
		cur_userid=rsyg("id")
	end if
	rsyg.close
	set rsyg = nothing
%>
<div id="loadingimg" align="center" style="width:100%; padding-top:100px; float:left; display:none"><img src="../Image/loading.gif" width="16" height="16"><br>
  <br>
<div id="loadingtext">������������,���Ե�...</div></div>
<script language="javascript">loadingShow();</script>
<%
Response.Flush()%>
</p>
<div align="center" class="style6">
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center"><%
response.write datearea		%><span class="STYLE7"> [
          <%if typed=1 then response.Write "����"
	if typed=2 then response.Write "����ʦ"
	if typed=4 then response.Write "��Ӱʦ"
	if typed=5 then response.Write GetDutyName(5)
	if typed=12 then response.Write "��Ӱʦ����"
	if typed=11 then response.Write "��ɴ����Ա"
	if typed=14 then response.Write GetDutyName(14)
	response.Write ":"&cur_peplename
	%>
          ]</span>
          <%
	if typed=4 then
		response.write "ѡƬҵ����"
	elseif typed=5 then
		response.write "ҵ������"
		'fujia2_save_money= conn.execute("select sum(money) from save_money where xiangmu_id in (select xiangmu_id from fujia2 where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id&") and type=3")(0)
'		goumai_save_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and xiangmu_id in (select xiangmu_id from goumai where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id&") and type=4")(0)
'		goumaijilu_save_money=conn.execute("select sum(money) from goumai_jilu where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')")(0)
'		response.write "&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>(���ջ�ױ:"&formatnumber(fujia2_save_money,1,0,0,0)&"Ԫ&nbsp;&nbsp;��黯ױ:"&formatnumber(goumai_save_money,1,0,0,0)&"Ԫ&nbsp;&nbsp;��ɢ����:"&formatnumber(goumaijilu_save_money,1,0,0,0)&"Ԫ)</font>"
		
	else
		
		if typed=1 then
			response.write "���½ӵ�ҵ������"
			savemoney = 0
			savemoney1 = 0
			savemoney2 = 0
			hq_savemoney = 0
			ls_money = 0
			
			'��ϵ
			set rstx = conn.execute("select s.id,s.money,s.beizhu,x.userid,x.userid2,x.userid3 from save_money s inner join shejixiadan x on s.xiangmu_id=x.id where "&GetSqlCheckDateString("s.times")&" and (x.userid='"&userid&"' or x.userid2='"&userid&"' or x.userid3='"&userid&"') and [s.type]=1")
			do while not rstx.eof
				if not isnull(rstx("userid2")) and rstx("userid2")<>"" and not isnull(rstx("userid3")) and rstx("userid3")<>"" then 
					count111=3
				elseif (not isnull(rstx("userid2")) and rstx("userid2")<>"") or (not isnull(rstx("userid3")) and rstx("userid3")<>"") then
					count111=2
				else
					count111=1
				end if
				savemoney = savemoney + rstx("money")/count111
				if rstx("beizhu")="���𸶿�" or rstx("beizhu")="���𸶿�" then
					savemoney1 = savemoney1 + rstx("money")/count111
				else
					savemoney2 = savemoney2 + rstx("money")/count111
				end if
				rstx.movenext
			loop
			rstx.close
			set rstx=nothing
			
			'����
			set rshq = conn.execute("select * from save_money where "&GetSqlCheckDateString("times")&" and xiangmu_id in (select id from shejixiadan where (ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"') and [type]=2)")
			do while not rshq.eof
				set rsxdx = conn.execute("select ky_name,ky_name2 from shejixiadan where id="&rshq("xiangmu_id"))
				if not rsxdx.eof then
					if not isnull(rsxdx("ky_name2")) and rsxdx("ky_name2")<>"" then 
						count222=2
					else
						count222=1
					end if
					hq_savemoney = hq_savemoney + rshq("money")/count222
				end if
				rsxdx.close
				set rsxdx = nothing
				rshq.movenext
			loop
			rshq.close
			set rshq=nothing
			
			'savemoney=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and xiangmu_id in (select id from shejixiadan where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and [type]=1)")(0)
			'if isnull(savemoney) then savemoney=0
			
			'hq_savemoney=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and xiangmu_id in (select id from shejixiadan where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and [type]=2)")(0)
			'if isnull(hq_savemoney) then hq_savemoney=0
			
			ls_money=conn.execute("select sum(money) from goumai_jilu where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&GetSqlCheckDateString("times")&"")(0)
			if isnull(ls_money) then ls_money=0
			
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>��ϵ�ɿ�:"&formatnumber(savemoney,1,0,0,0)&"Ԫ(����:"&formatnumber(savemoney1,1,0,0,0)&"Ԫ&nbsp;���ս�:"&formatnumber(savemoney2,1,0,0,0)&"Ԫ)&nbsp;&nbsp;���ڽɿ�:"&formatnumber(hq_savemoney,1,0,0,0)&"Ԫ&nbsp;&nbsp;��ɢ����:"&formatnumber(ls_money,1,0,0,0)&"Ԫ</font>"
			'response.write allmoney&"-"&savemoney
		else
			response.write "ҵ������"
		end if
	end if
	%> <br>
        </div></td>
      </tr>
  </table>
</div>

<%
select case typed
case 1
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&GetSqlCheckDateString("times"),conn,1,1

Call init_key()
Call InitConsInfo()  'ǰ�ڽ��ֶβ���
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" style="richness:1px">
  <tr bgcolor="#99FFFF">
    <td width="60" height="19" align="center">����</td>
    <td width="80" align="center">�ͻ�</td>
    <td align="center">��ϵ����</td>
    <td align="center">��ϵ����</td>
    <td width="70" align="center">����ϵ��</td>
    <td width="200" align="center">��ϵ�ɷ�/(�Ŷ�)/
    ����/���ս�</td>
    <td width="120" align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font>/����</td>
    <td width="70" align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
    <td width="70" align="center">��ϵǷ��</td>
    <td width="70" align="center">����</td>
    <td align="center">��ע</td>
  </tr>
  <%
  do while not rs.eof
  str_sm=""
  	count111=1
	if not isnull(rs("userid2")) and rs("userid2")<>"" then count111=count111+1
  	if not isnull(rs("userid3")) and rs("userid3")<>"" then count111=count111+1
	
  MonthWedsuitCost = MonthWedsuitCost + getWedsuitCost(rs("id"))
	
  bk_jixiang=0
  bk_fujia=0
  
  '�������½ɺ��ڿ�
  hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
  	if isnull(money2) then money2=0
	count222 = 1
	if rs("ky_name2")<>"" and not isnull(rs("ky_name2")) then
		count222 = 2
	end if
	sm2_money=money2
	hq_indate_savemoney=hq_indate_savemoney/count222
  
  '�����ܺ���
  hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
  if isnull(hq_money) then hq_money = 0
  
  '�����ܺ��ڽɿ�
  hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
  
  'hq_minesavemoney = conn.execute("select sum(money) from save_money where [type]=2 and userid='"&userid&"' and xiangmu_id="&rs("id"))(0)
  
  set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
  do while not rshq.eof
  	if rshq("userid")=userid or rshq("userid2")=userid then
	  if rshq("userid")<>"" and not isnull(rshq("userid2")) then
		hq_mymoney = hq_mymoney + rshq("money")/2
	  else
	  	hq_mymoney = hq_mymoney + rshq("money")
  	  end if
	end if
	rshq.movenext
  loop
  rshq.close
  set rshq=nothing
  
  if isnull(hq_savemoney) then hq_savemoney = 0
  
  money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
  if isnull(money1) then money1=0
  
  '��Ƿ��
  hq_notsavemoney=hq_notsavemoney+hq_money-hq_savemoney
  
  '�ܺ���
  hq_allmoney=hq_allmoney+hq_money
  
  '�����ܺ��ڽɿ�
  hq_indate_allsavemoney=hq_indate_allsavemoney+hq_indate_savemoney
  customerlosttypeid = GetFieldDataBySQL("select customerlosttype from kehu where id="&rs("kehu_id"),"int",0)

  For aai = 1 To UBound(arr_cons_info, 1)
	If CInt(arr_cons_info(aai, 0))=customerlosttypeid Then 
	  arr2 = arr_cons_info(aai, 3)
	  arr3 = arr_cons_info(aai, 4)
	  arr4 = arr_cons_info(aai, 5)
	  If IsArray(arr2) Then 
		  for cci = 1 to ubound(arr2)
			if cint(arr2(cci))>=money1 then
				arr3(cci) = arr3(cci) + money1/count111
				arr4(cci) = arr4(cci) + 1
				exit for
			end if
		  Next
		  arr_cons_info(aai, 4) = arr3
		  arr_cons_info(aai, 5) = arr4
		  Exit For 
	  End if
	End If 
  Next 
  %>
  <tr bgcolor="#FFFFFF" id="<%="tr_"&rs("id")%>">
    <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		msidlist=msidlist&rs("id")&","
	%>    </td>
    <td align="center">
    <%
	 response.Write conn.execute("select lxpeple from kehu where id="&rs("kehu_id"))(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>  
    <td align="center"><%=GetFieldDataBySQL("select CustomerLostType.title from CustomerLostType inner join kehu on CustomerLostType.id=kehu.CustomerLostType where kehu.id="&rs("kehu_id"),"str","&nbsp;")%></td>
    <td align="center"><%=GetFieldDataBySQL("select jixiang from jixiang where id="&rs("jixiang"),"str","&nbsp;")%></td>
    <td align="center"><% 
		jx_money = rs("jixiang_money")
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
    <td align="center">
  <%
	sm1_money=money1/count111
	if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then response.Write formatnumber(sm1_money,1,0,0,0)
	if not isnull(rs("userid")) and rs("userid")<>"" and rs("userid")<>userid then str_sm=str_sm&"/"&GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid")&"'","str","N/A")
	if not isnull(rs("userid2")) and rs("userid2")<>"" and rs("userid2")<>userid then str_sm=str_sm&"/"&GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid2")&"'","str","N/A")
	if not isnull(rs("userid3")) and rs("userid3")<>"" and rs("userid3")<>userid then str_sm=str_sm&"/"&GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid3")&"'","str","N/A")
	if left(str_sm,1)="/" then response.write mid(str_sm,2)
	
	dim dd_dingjin,dd_paizhao
	dd_dingjin=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and (beizhu='���𸶿�' or beizhu='���𸶿�')")(0)
	if isnull(dd_dingjin) then dd_dingjin=0
	dd_paizhao=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and beizhu<>'���𸶿�' and beizhu<>'���𸶿�'")(0)
	if isnull(dd_paizhao) then dd_paizhao=0
	response.write "/"&dd_dingjin&"/"&dd_paizhao
	%></td>
    <td align="center"><%
	money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and xiangmu_id in (select id from shejixiadan where ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"')")(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
	if rs("ky_name")<>cur_peplename then
			response.Write "/"&rs("ky_name")
	  end if
	  if rs("ky_name2")<>cur_peplename then
			response.Write "/"&rs("ky_name2")
	  end if
	%></td>
    <td align="center"><%=FinalMoneySum(rs("id"),False)%></td>
    <td align="center"><%=jx_money-money1%></td>
    <td align="center"><%=getPerStep(rs("id"))%>&nbsp;</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
  
  if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then
  	jixiang_money = jixiang_money + jx_money
  	money00=money00+money1
	tx_savemoney = conn.execute("select sum([money]) from save_money where [type]=1 and xiangmu_id="&rs("id"))(0)
  	if isnull(tx_savemoney) then tx_savemoney=0
  	if tx_savemoney=rs("jixiang_money") and conn.execute("select count(*) from save_money where xiangmu_id="&rs("id"))(0)>0 then
  		ReceivablesMoney = ReceivablesMoney + (rs("jixiang_money")/count111)
  	end if
  end if
  if hq_money=hq_indate_savemoney then 
  	RecFujiaMoney = RecFujiaMoney+hq_mymoney
	AllRecFujiaMoney = AllRecFujiaMoney+hq_money
  end if
  rs.movenext
  i=i+1
loop
rs.close
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;�ӵ���ϵ�ܽ��
      <%=formatnumber(jixiang_money,1,0,0,0)%>
Ԫ&nbsp; <%if session("level")=10 or (session("level")=7 and session("zhuguan")=1) then
		dim arr_cb,stime,etime
		if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
			stime = yeard&"-"&monthd&"-1"
			etime = cstr(DateAdd("d",-1,DateAdd("m",1,cdate(stime))))
		end if
		if fromtime<>"" and totime<>"" then
			stime = fromtime
			etime = totime
		end if
		arr_cb = GetCostCalcuation(stime,etime,userid,"",msidlist,"0","userid,userid2,userid3")
		%>
 ��ϵ�ɱ� <%=arr_cb(0,1)%> Ԫ<%end if%>&nbsp;&nbsp;��ϵδ�� 
<%
	response.write formatnumber(jixiang_money-money00,1,0,0,0)
'	jixiang_choucheng=money11*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
'	'response.write formatnumber(money11,1,0,0,0)
'	if isnull(jixiang_choucheng) then jixiang_choucheng=0
'	response.write formatnumber(jixiang_money-money11,1,0,0,0)%> 
Ԫ&nbsp;&nbsp;<span class="STYLE11">������ϵδǷ���ܽ�� <%=formatnumber(ReceivablesMoney,1,0,0,0)%>&nbsp;Ԫ������ǰ�տ</span><br><%if IsArray(arr_cons_info) then
	response.write "&nbsp;ǰ�ڽ��ֶκϼ�<br>"
	For aai = 1 To UBound(arr_cons_info, 1)
		arr1 = arr_cons_info(aai, 2)
		arr2 = arr_cons_info(aai, 3)
		arr3 = arr_cons_info(aai, 4)
		arr4 = arr_cons_info(aai, 5)
		If IsArray(arr1) Then 
			response.write "&nbsp;" & arr_cons_info(aai, 1) & "��"
			for cci = 1 to ubound(arr1)
				response.write arr1(cci)&" ~ "&arr2(cci)&"Ԫ("&arr4(cci)&")��"& Formatnumber(arr3(cci),1,0,0,0)&"Ԫ&nbsp;&nbsp;&nbsp;"
			Next
			response.write "<br>"
		End if
	Next 

	'for cci = 1 to ubound(arr_cons_maxmoney)
	'	response.write arr_cons_minmoney(cci)&" ~ "&arr_cons_maxmoney(cci)&"Ԫ��"& Formatnumber(arr_cons_vol(cci),1,0,0,0)&"Ԫ&nbsp;&nbsp;&nbsp;"
	'next
	
end if%>&nbsp;<%
		ds1_all = conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where  (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times"))(0)
		ds2_all = conn.execute("select sum(s.jixiang_money) from shejixiadan s inner join kehu k on s.kehu_id=k.id where (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times"))(0)
		if isnull(ds1_all) then ds1_all=0
		if isnull(ds2_all) then ds2_all=0
		
		ds1_count=0
		ds2_count=0
		set rslost = conn.execute("select * from CustomerLostType order by px")
		do while not rslost.eof
			ds1 = conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.CustomerLostType="&rslost("id")&" and (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times"))(0)
			ds2 = conn.execute("select sum(s.jixiang_money) from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.CustomerLostType="&rslost("id")&" and (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times"))(0)
			if isnull(ds1) then ds1=0
			if isnull(ds2) then ds2=0
			ds1_count = ds1_count + ds1
			ds2_count = ds2_count + ds2
			response.write rslost("title")&"ƽ����� "
			if ds1=0 then 
				response.write ".0"
			else
				response.write formatnumber(ds2/ds1,1,0,0,0)
			end if
			response.write " Ԫ&nbsp;&nbsp;&nbsp;"
			rslost.movenext
		loop
		rslost.close
		set rslost = nothing
		response.write "����ƽ����� "
		if ds1_all-ds1_count=0 then 
			response.write ".0"
		else
			response.write formatnumber((ds2_all-ds2_count)/(ds1_all-ds1_count),1,0,0,0)
		end if
		response.write " Ԫ"
%><br>
&nbsp;��������Ӱ
      <%
		sycount=0
		syall=conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times")&" and not isnull(s.lc_cp)")(0)
		if isnull(syall) then syall=0
		response.write syall
%>�� (<%set rssy = conn.execute("select * from CustomerLostType order by px")
		do while not rssy.eof
			sy = conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.CustomerLostType="&rssy("id")&" and (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times")&" and not isnull(s.lc_cp)")(0)
			if isnull(sy) then sy=0
			sycount = sycount + sy
			response.write rssy("title")&sy&",&nbsp;"
			rssy.movenext
		loop
		rssy.close
		set rssy = nothing
		response.write "����" & syall - sycount
%>)<br>
&nbsp;����δ��Ӱ
<%
		sy=conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times")&" and isnull(s.lc_cp)")(0)
		if isnull(sy) then sy=0
		response.write sy
%>�� (<%set rssy = conn.execute("select * from CustomerLostType order by px")
		do while not rssy.eof
			sy = conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and k.CustomerLostType="&rssy("id")&" and "&GetSqlCheckDateString("s.times")&" and isnull(s.lc_cp)")(0)
			if isnull(sy) then sy=0
			response.write rssy("title")&sy
			rssy.movenext
			if not rssy.eof then response.write ",&nbsp;"
		loop
		rssy.close
		set rssy = nothing%>)
<%if request.form("basepoint_flag")="show" then%>&nbsp;&nbsp; ������ϵƽ������: 
<%if MonthWedsuitCost=0 or jixiang_money=0 then
	response.write "0"
else
	response.write formatnumber(jixiang_money / MonthWedsuitCost,3) * 100 & " %"
end if
%>&nbsp;&nbsp; ����ϵƽ������: 
<%
	set rs_allxm = conn.execute("select id,jixiang_money from shejixiadan where userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"' order by id")
	do while not rs_allxm.eof
		AllXiangmuMoney = AllXiangmuMoney + rs_allxm("jixiang_money")
		AllWedsuitCost = AllWedsuitCost + getWedsuitCost(rs_allxm("id"))
		rs_allxm.movenext
	loop
	rs_allxm.close
	set rs_allxm = nothing
	if AllWedsuitCost=0 or AllXiangmuMoney=0 then
		response.write "0"
	else
		response.write formatnumber(AllXiangmuMoney / AllWedsuitCost,3) * 100 & " %"
	end if
end if
%>
<br>
&nbsp;���¿ͻ��ɽ�
      <%
		dscount=0
		dsall=conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times"))(0)
		if isnull(dsall) then dsall=0
		response.write dsall
%>�� (<%set rslost = conn.execute("select * from CustomerLostType order by px")
		do while not rslost.eof
			ds = conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.CustomerLostType="&rslost("id")&" and (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"') and "&GetSqlCheckDateString("s.times"))(0)
			if isnull(ds) then ds=0
			dscount = dscount + ds
			response.write rslost("title")&ds&",&nbsp;"
			rslost.movenext
		loop
		rslost.close
		set rslost = nothing
		response.write "����" & dsall - dscount
%>)<br>
&nbsp;���¿ͻ���ʧ
<%
		ds=conn.execute("select count(*) from kehu where islost=1 and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&GetSqlCheckDateString("times"))(0)
		if isnull(ds) then ds=0
		response.write ds
%>�� (<%set rslost = conn.execute("select * from CustomerLostType order by px")
		do while not rslost.eof
			ds = conn.execute("select count(*) from kehu where islost=1 and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and  CustomerLostType="&rslost("id")&" and "&GetSqlCheckDateString("times"))(0)
			if isnull(ds) then ds=0
			response.write rslost("title")&ds
			rslost.movenext
			if not rslost.eof then response.write ",&nbsp;"
		loop
		rslost.close
		set rslost = nothing%>)</td>
  </tr>
</table>

<br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" align="center"><%response.write datearea%>
      &nbsp;ѡƬ��ϸ��</td>
  </tr>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" style="richness:1px">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">����ϵ��</td>
    <td align="center">��ϵ�ɷ�/(�Ŷ�)</td>
    <td align="center">ѡƬ�����ܽ��</td>
    <td align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font>/����</td>
    <td width="16%" align="center">��Ƭ���͡�</td>
    <td align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
  </tr>
  <%
  Call init_key()
  Call InitConsInfo()'���ڽ��ֶβ���

  rs.open "select * from shejixiadan where (ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"') and "&GetSqlCheckDateString("lc_ky"),conn,1,1

  do while not rs.eof
  	str_sm=""
  	count111=1
	if not isnull(rs("userid2")) and rs("userid2")<>"" then count111 = count111+1
	if not isnull(rs("userid3")) and rs("userid3")<>"" then count111 = count111+1
	
   	count222 = 1
	if rs("ky_name2")<>"" and not isnull(rs("ky_name2")) then count222 = 2
  
  	'�������½ɺ��ڿ�
  	hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  	if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
  
  	'���ڽ��ֶ�
	customerlosttypeid = GetFieldDataBySQL("select customerlosttype from kehu where id="&rs("kehu_id"),"int",0)
	For aai = 1 To UBound(arr_cons_info, 1)
		If CInt(arr_cons_info(aai, 0))=customerlosttypeid Then 
		  arr2 = arr_cons_info(aai, 3)
		  arr3 = arr_cons_info(aai, 4)
		  arr4 = arr_cons_info(aai, 5)
		  If IsArray(arr2) Then 
			  for cci = 1 to ubound(arr2)
				if cint(arr2(cci))>=hq_indate_savemoney then
					arr3(cci) = arr3(cci) + hq_indate_savemoney/count222
					arr4(cci) = arr4(cci) + 1
					exit for
				end if
			  Next
			  arr_cons_info(aai, 4) = arr3
			  arr_cons_info(aai, 5) = arr4
			  Exit For 
		  End if
		End If 
	Next 
  
  	if isnull(money2) then money2=0
	sm2_money=money2
	hq_indate_savemoney=hq_indate_savemoney/count222
  
  	'�����ܺ���
  	hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
  	if isnull(hq_money) then hq_money = 0
  
  	'�����ܺ��ڽɿ�
  	hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
  
  
  	if hq_money=hq_savemoney then
  		ReceivablesMoney = ReceivablesMoney + hq_money
  	end if

  	'if hq_money=hq_indate_savemoney then 
  	'	RecFujiaMoney = RecFujiaMoney+hq_mymoney
	'AllRecFujiaMoney = AllRecFujiaMoney+hq_money
  	'end if
  
  	set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
  	do while not rshq.eof
		hq_minemoney = hq_minemoney + rshq("money") / count222
		rshq.movenext
  	loop
  	rshq.close
  	set rshq=nothing
  
  tmp_money = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
  if isnull(tmp_money) then tmp_money = 0
  hq_minesavemoney = hq_minesavemoney + tmp_money / count222
  
  if isnull(hq_savemoney) then hq_savemoney = 0
  
  '��Ƿ��
  hq_notsavemoney=hq_notsavemoney+hq_money-hq_savemoney
  
  '�ܺ���
  'hq_allmoney=hq_allmoney+hq_money
  
  '�����ܺ��ڽɿ�
  hq_indate_allsavemoney=hq_indate_allsavemoney+hq_indate_savemoney
  
  hqmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  if isnull(hqmoney) then hqmoney=0
	
  
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		msidlist=msidlist&rs("id")&","
	%>    </td>
    <td align="center"><%
	 response.Write conn.execute("select lxpeple from kehu where id="&rs("kehu_id"))(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>
    <td align="center"><% 
		jx_money = rs("jixiang_money")/count111
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
    <td align="center"><%money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money1) then money1=0
	if rs("userid")<>userid and rs("userid2")<>userid and rs("userid3")<>userid then money1=0
	sm1_money=money1/count111
	if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then response.Write formatnumber(sm1_money,1,0,0,0)
	if rs("userid")<>"" and rs("userid")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)
	if rs("userid2")<>"" and rs("userid2")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid2")&"'")(0)
	if rs("userid3")<>"" and rs("userid3")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid3")&"'")(0)
	if left(str_sm,1)="/" then response.write mid(str_sm,2)
	%></td>
    <td align="center"><%
	response.write Formatnumber(hqmoney/count222,1,0,0,0)
	hq_allmoney = hq_allmoney + hqmoney/count222
	%></td>
    <td align="center" bgcolor="#ffffff"><%
	money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and xiangmu_id in (select id from shejixiadan where ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"')")(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
	if rs("ky_name")<>cur_peplename then
			response.Write "/"&rs("ky_name")
	  end if
	  if rs("ky_name2")<>cur_peplename then
			response.Write "/"&rs("ky_name2")
	  end if
	%></td>
    <td align="center" bgcolor="#ffffff"><%if rs("ky_name")<>cur_peplename and rs("ky_name2")<>cur_peplename then
		response.write "0"
	else%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
        <tr>
          <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
          <td>&nbsp;<%=rsdg("all_sl")%></td>
        </tr>
        <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
      </table>
    <%end if%></td>
    <td align="center"><%=FinalMoneySum(rs("id"),False)%></td>
  </tr>
  <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
 
  rs.movenext
  i=i+1
loop
rs.close()
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;�����ܽ�� <%=Formatnumber(hq_allmoney,1,0,0,0)%> Ԫ&nbsp;&nbsp;&nbsp;<%if session("level")=10 or (session("level")=7 and session("zhuguan")=1) then
		if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
			stime = yeard&"-"&monthd&"-1"
			etime = cstr(DateAdd("d",-1,DateAdd("m",1,cdate(stime))))
		end if
		if fromtime<>"" and totime<>"" then
			stime = fromtime
			etime = totime
		end if
		arr_cb = GetCostCalcuation(stime,etime,userid,"",msidlist,"1","ky_name,ky_name2")
		%>
 ���ڳɱ� <%=arr_cb(1,1)%> Ԫ<%end if%>&nbsp;&nbsp;�ѽ� <%'=hq_allmoney-hq_notsavemoney
	if isnull(hq_minesavemoney) then hq_minesavemoney=0
	response.write hq_minesavemoney%> Ԫ &nbsp;&nbsp;δ�� <%'=hq_notsavemoney
	response.write hq_minemoney-hq_minesavemoney%> Ԫ&nbsp;&nbsp; ������ڿ� <%=ReceivablesMoney%> Ԫ<br>&nbsp;<%
		set rs_ds1 = server.createobject("adodb.recordset")
		set rs_ds2 = server.createobject("adodb.recordset")
		set rs_ds3 = server.createobject("adodb.recordset")
		
		rs_ds1.open "select distinct s.id from shejixiadan s inner join kehu k on s.kehu_id=k.id where  (s.ky_name='"&peplename&"' or s.ky_name2='"&peplename&"') and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		if not (rs_ds1.eof and rs_ds1.bof) then
			ds1_all = rs_ds1.recordcount
		else
			ds1_all = 0
		end if
		rs_ds1.close
		
		rs_ds3.open "select distinct s.id from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where (s.ky_name='"&peplename&"' or s.ky_name2='"&peplename&"') and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		if not (rs_ds3.eof and rs_ds3.bof) then
			ds3_all = rs_ds3.recordcount
		else
			ds3_all = 0
		end if
		rs_ds3.close
		
		ds2_all = 0
		rs_ds2.open "select s.ky_name,s.ky_name2,f.money from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where (s.ky_name='"&peplename&"' or s.ky_name2='"&peplename&"') and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		do while not rs_ds2.eof
			if not isnull(rs_ds2("ky_name2")) and rs_ds2("ky_name2")<>"" then
				ds2_all = ds2_all + rs_ds2("money")/2
			else
				ds2_all = ds2_all + rs_ds2("money")
			end if
			rs_ds2.movenext
		loop
		rs_ds2.close
		
		ds_count=0		'����
		ds1_count=0		'ѡƬ��¼����
		ds2_count=0		'ѡƬ���Ѻϼ�
		ds3_count=0		'ѡ�����Ѽ�¼����
		set rslost = conn.execute("select * from CustomerLostType order by px")
		do while not rslost.eof
			ds1 = 0
			ds2 = 0
			ds3 = 0
			
			rs_ds1.open "select distinct s.id from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.CustomerLostType="&rslost("id")&" and (s.ky_name='"&peplename&"' or s.ky_name2='"&peplename&"') and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			if not (rs_ds1.eof and rs_ds1.bof) then
				ds1 = rs_ds1.recordcount
			else
				ds1 = 0
			end if
			rs_ds1.close
			
			rs_ds3.open "select distinct s.id from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where k.CustomerLostType="&rslost("id")&" and (s.ky_name='"&peplename&"' or s.ky_name2='"&peplename&"') and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			if not (rs_ds3.eof and rs_ds3.bof) then
				ds3 = rs_ds3.recordcount
			else
				ds3 = 0
			end if
			rs_ds3.close
			
			rs_ds2.open "select s.ky_name,s.ky_name2,f.money from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where k.CustomerLostType="&rslost("id")&" and (s.ky_name='"&peplename&"' or s.ky_name2='"&peplename&"') and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			do while not rs_ds2.eof
				if not isnull(rs_ds2("ky_name2")) and rs_ds2("ky_name2")<>"" then
					ds2 = ds2 + rs_ds2("money")/2
				else
					ds2 = ds2 + rs_ds2("money")
				end if
				rs_ds2.movenext
			loop
			rs_ds2.close
			
			ds1_count = ds1_count + ds1
			ds2_count = ds2_count + ds2
			ds3_count = ds3_count + ds3
			response.write rslost("title")&"ѡƬ"&ds1&"�� "
			response.write "δ����"& ds1-ds3 &"�� "
			response.write "��"&ds2&"Ԫ ƽ�����"
			if ds1=0 then 
				response.write ".0"
			else
				response.write formatnumber(ds2/ds1,1,0,0,0)
			end if
			response.write " Ԫ&nbsp;&nbsp;&nbsp;"
			ds_count = ds_count + 1
			if ds_count mod 2 = 0 then response.write "<br>&nbsp;"
			rslost.movenext
		loop
		rslost.close
		set rslost = nothing
		
		response.write "����ѡƬ"&ds1_all-ds1_count&"�� "
		response.write "δ����"& (ds1_all-ds3_all)-(ds1_count-ds3_count) &"�� "
		response.write "��"& ds2_all-ds2_count &"Ԫ ƽ�����"
		if (ds1_all-ds3_all)-(ds1_count-ds3_count)=0 then 
			response.write ".0"
		else
			response.write formatnumber((ds2_all-ds2_count)/(ds1_all-ds1_count),1,0,0,0)
		end if
		response.write " Ԫ"
%><br><%if IsArray(arr_cons_info) And IsArray(arr_cons_info(1, 2)) then
	response.write "&nbsp;���ڽ��ֶ�ͳ��<br>"
	For aai = 1 To UBound(arr_cons_info, 1)
		arr1 = arr_cons_info(aai, 2)
		arr2 = arr_cons_info(aai, 3)
		arr3 = arr_cons_info(aai, 4)
		arr4 = arr_cons_info(aai, 5)
		If IsArray(arr1) Then 
			response.write "&nbsp;" & arr_cons_info(aai, 1) & "��"
			for cci = 1 to ubound(arr1)
				response.write arr1(cci)&" ~ "&arr2(cci)&"Ԫ("&arr4(cci)&")��"& Formatnumber(arr3(cci),1,0,0,0)&"Ԫ&nbsp;&nbsp;&nbsp;"
			Next
			response.write "<br>"
		End if
	Next 
End if%>&nbsp;���º��ڻ������� <%
	set rstemp = conn.execute("SELECT sum(y.in_money * f.sl) FROM (shejixiadan s INNER JOIN fujia f ON s.ID = f.xiangmu_id) INNER JOIN yunyong y ON f.jixiang = y.ID where y.in_money<>0 and y.type=1 and (s.ky_name='"&peplename&"' or s.ky_name2='"&peplename&"') and "&GetSqlCheckDateString("s.lc_ky")&" group by f.jixiang")
	do while not rstemp.eof
		MonthFujiaCost = MonthFujiaCost + rstemp(0)
		rstemp.movenext
	loop
	rstemp.close
	set rstemp = nothing
	if isnull(MonthFujiaCost) or trim(cstr(MonthFujiaCost))="" then MonthFujiaCost=0
	if hq_allmoney=0 or MonthFujiaCost=0 then
		response.write "0"
	else
		response.write formatnumber(MonthFujiaCost/hq_allmoney,3) * 100 & " %"
	end if
	%>&nbsp;&nbsp;&nbsp;�ܺ��ڻ������� <%
	set rstemp = conn.execute("SELECT sum(f.money), sum(y.in_money * f.sl) FROM (shejixiadan s INNER JOIN fujia f ON s.ID = f.xiangmu_id) INNER JOIN yunyong y ON f.jixiang = y.ID where y.in_money<>0 and y.type=1 and (s.ky_name='"&peplename&"' or s.ky_name2='"&peplename&"') group by f.jixiang")
	do while not rstemp.eof
		AllFujiaMoney = AllFujiaMoney + rstemp(0)
		AllFujiaCost = AllFujiaCost + rstemp(1)
		rstemp.movenext
	loop
	rstemp.close
	set rstemp = nothing
	if isnull(AllFujiaMoney) or trim(cstr(AllFujiaMoney))="" then AllFujiaMoney=0
	if isnull(AllFujiaCost) or trim(cstr(AllFujiaCost))="" then AllFujiaCost=0
	if AllFujiaMoney=0 or AllFujiaCost=0 then
		response.write "0"
	else
		response.write formatnumber(AllFujiaMoney/AllFujiaCost,3) * 100 & " %"
	end if
	%></td>
  </tr>
</table>
<br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" align="center"><%response.write datearea%>
      &nbsp;��ϵ������ϸ��</td>
  </tr>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" style="richness:1px">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">����ϵ��</td>
    <td align="center">��ϵ�ɷ�/(�Ŷ�)</td>
    <td width="16%" align="center">ѡƬ�����ܽ��</td>
    <td align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font>/����</td>
    <td align="center">��Ƭ���͡�</td>
    <td align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
  </tr>
  <%
  Call init_key()
  Call InitConsInfo()  '��ϵ����ֶβ���

  rs.open "select * from shejixiadan where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&sql_time&" and id in (select xiangmu_id from save_money where [type]=1 and "&GetSqlCheckDateString("times")&")",conn,1,1

  'msidlist=","
  do while not rs.eof
  str_sm=""
  if not isnull(rs("userid3")) and rs("userid3")<>"" then 
	count111=3
	elseif not isnull(rs("userid2")) and rs("userid2")<>"" then
	count111=2
	else
	count111=1
	end if
  
  if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then
  	jixiang_money = jixiang_money + jx_money
  	money00=money00+money1
	tx_savemoney = conn.execute("select sum([money]) from save_money where [type]=1 and xiangmu_id="&rs("id"))(0)
  	if isnull(tx_savemoney) then tx_savemoney=0
  	if tx_savemoney=rs("jixiang_money") and conn.execute("select count(*) from save_money where xiangmu_id="&rs("id"))(0)>0 then
  		ReceivablesMoney = ReceivablesMoney + (rs("jixiang_money")/count111)
  	end if
  end if
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		msidlist=msidlist&rs("id")&","
	%>
    </td>
    <td align="center"><%
	 response.Write conn.execute("select lxpeple from kehu where id="&rs("kehu_id"))(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>
    <td align="center"><% 
		jx_money = rs("jixiang_money")/count111
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
    <td align="center"><%money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money1) then money1=0
	if rs("userid")<>userid and rs("userid2")<>userid and rs("userid3")<>userid then money1=0
	sm1_money=money1/count111

	customerlosttypeid = GetFieldDataBySQL("select customerlosttype from kehu where id="&rs("kehu_id"),"int",0)
    For aai = 1 To UBound(arr_cons_info, 1)
		If CInt(arr_cons_info(aai, 0))=customerlosttypeid Then 
		  arr2 = arr_cons_info(aai, 3)
		  arr3 = arr_cons_info(aai, 4)
          arr4 = arr_cons_info(aai, 5)
		  If IsArray(arr2) Then 
			  for cci = 1 to ubound(arr2)
				if cint(arr2(cci))>=money1 Then
					arr3(cci) = arr3(cci) + sm1_money
					arr4(cci) = arr4(cci) + 1
					exit for
				end if
			  Next
			  arr_cons_info(aai, 4) = arr3
			  arr_cons_info(aai, 5) = arr4
			  Exit For 
		  End if
		End If 
	Next 

	bk_jixiang = bk_jixiang + money1
	if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then response.Write formatnumber(sm1_money,1,0,0,0)
	if rs("userid")<>"" and rs("userid")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)
	if rs("userid2")<>"" and rs("userid2")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid2")&"'")(0)
	if rs("userid3")<>"" and rs("userid3")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid3")&"'")(0)
	if left(str_sm,1)="/" then response.write mid(str_sm,2)
	%></td>
    <td align="center" bgcolor="#ffffff"><%
	hqallmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	if isnull(hqallmoney) then hqallmoney=0
	response.write Formatnumber(hqallmoney/count222,1,0,0,0)
	%></td>
    <td align="center"><%
	money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and xiangmu_id in (select id from shejixiadan where ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"')")(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
	if rs("ky_name")<>cur_peplename then
			response.Write "/"&rs("ky_name")
	  end if
	  if rs("ky_name2")<>cur_peplename then
			response.Write "/"&rs("ky_name2")
	  end if
	%></td>
    <td align="center" bgcolor="#ffffff"><%if rs("ky_name")<>cur_peplename and rs("ky_name2")<>cur_peplename then
		response.write "0"
	else%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
          <tr>
            <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
            <td>&nbsp;<%=rsdg("all_sl")%></td>
          </tr>
          <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
        </table>
      <%end if%></td>
    <td align="center"><%=FinalMoneySum(rs("id"),False)%></td>
  </tr>
  <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
 
  rs.movenext
  i=i+1
loop
rs.close()
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;��ϵ���� <%=Formatnumber(bk_jixiang,1,0,0,0)%> Ԫ&nbsp;&nbsp;&nbsp; &nbsp;�ۼ���ϵǷ�� <%
	tmp_jixiang_money = conn.execute("select sum(jixiang_money) from shejixiadan where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')")(0)
	tmp_save_money = conn.execute("select sum(m.money) from save_money m inner join shejixiadan s on m.xiangmu_id=s.id where m.type=1 and (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"')")(0)
	if isnull(tmp_jixiang_money) then tmp_jixiang_money = 0
	if isnull(tmp_save_money) then tmp_save_money = 0
	response.write Formatnumber(tmp_jixiang_money-tmp_save_money,1,0,0,0)%> Ԫ&nbsp;&nbsp;&nbsp; &nbsp;������ϵ <%=Formatnumber(ReceivablesMoney,1,0,0,0)%> Ԫ<br><%

	if IsArray(arr_cons_info) then
	response.write "&nbsp;��ϵ������ֶκϼ�<br>"
	For aai = 1 To UBound(arr_cons_info, 1)
		arr1 = arr_cons_info(aai, 2)
		arr2 = arr_cons_info(aai, 3)
		arr3 = arr_cons_info(aai, 4)
		arr4 = arr_cons_info(aai, 5)
		If IsArray(arr1) Then 
			response.write "&nbsp;" & arr_cons_info(aai, 1) & "��"
			for cci = 1 to ubound(arr1)
				response.write arr1(cci)&" ~ "&arr2(cci)&"Ԫ("&arr4(cci)&")��"& Formatnumber(arr3(cci),1,0,0,0)&"Ԫ&nbsp;&nbsp;&nbsp;"
			Next
			response.write "<br>"
		End if
	Next 
	end if%></td>
  </tr>
</table>
<br>
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="30" align="center"><%response.write datearea%>
        &nbsp;ѡƬ������ϸ��</td>
    </tr>
</table>
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" style="richness:1px">
    <tr bgcolor="#99FFFF">
      <td height="19" align="center">����</td>
      <td align="center">�ͻ�</td>
      <td align="center">����ϵ��</td>
      <td align="center">��ϵ�ɷ�/(�Ŷ�)</td>
      <td width="16%" align="center">ѡƬ�����ܽ��</td>
      <td align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font>/����</td>
      <td align="center">��Ƭ���͡�</td>
      <td align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
    </tr>
    <%
  Call init_key()
  Call InitConsInfo()  '���ڲ�����ֶβ���
  
  rs.open "select * from shejixiadan where (ky_name='"&peplename&"' or ky_name2='"&peplename&"' ) and "&sql_time&" and id in (select xiangmu_id from save_money where [type]=2 and "&GetSqlCheckDateString("times")&")",conn,1,1

  'msidlist=","
  do while not rs.eof
  str_sm=""
  if not isnull(rs("userid3")) and rs("userid3")<>"" then 
	count111=3
	elseif not isnull(rs("userid2")) and rs("userid2")<>"" then
	count111=2
	else
	count111=1
	end if
 
  
  '�������½ɺ��ڿ�
  hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
  	if isnull(money2) then money2=0
	count222 = 1
	if rs("ky_name2")<>"" and not isnull(rs("ky_name2")) then
		count222 = 2
	end if
	sm2_money=money2
	hq_indate_savemoney=hq_indate_savemoney/count222
  
  '�����ܺ���
  hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
  if isnull(hq_money) then hq_money = 0
  
  '�����ܺ��ڽɿ�
  hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
  
  
  	if hq_money=hq_savemoney then
  		ReceivablesMoney = ReceivablesMoney + hq_money
  	end if

  'if hq_money=hq_indate_savemoney then 
  '	RecFujiaMoney = RecFujiaMoney+hq_mymoney
	'AllRecFujiaMoney = AllRecFujiaMoney+hq_money
  'end if
  
  'hq_minesavemoney = conn.execute("select sum(money) from save_money where [type]=2 and userid='"&userid&"' and xiangmu_id="&rs("id"))(0)
  set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
  do while not rshq.eof
  	if rshq("userid")=userid or rshq("userid2")=userid then
	  if rshq("userid")<>"" and not isnull(rshq("userid2")) then
		hq_mymoney = hq_mymoney + rshq("money")/2
	  else
	  	hq_mymoney = hq_mymoney + rshq("money")
  	  end if
	end if
	rshq.movenext
  loop
  rshq.close
  set rshq=nothing
  
  if isnull(hq_savemoney) then hq_savemoney = 0
  
  '��Ƿ��
  hq_notsavemoney=hq_notsavemoney+hq_money-hq_savemoney
  
  '�ܺ���
  hq_allmoney=hq_allmoney+hq_money
  
  '�����ܺ��ڽɿ�
  hq_indate_allsavemoney=hq_indate_allsavemoney+hq_indate_savemoney
  %>
    <tr bgcolor="#FFFFFF">
      <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		msidlist=msidlist&rs("id")&","
	%>      </td>
      <td align="center"><%
	 response.Write conn.execute("select lxpeple from kehu where id="&rs("kehu_id"))(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>
      <td align="center"><% 
		jx_money = rs("jixiang_money")/count111
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
      <td align="center"><%money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money1) then money1=0
	if rs("userid")<>userid and rs("userid2")<>userid and rs("userid3")<>userid then money1=0
	sm1_money=money1/count111
	if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then response.Write formatnumber(sm1_money,1,0,0,0)
	if rs("userid")<>"" and rs("userid")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)
	if rs("userid2")<>"" and rs("userid2")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid2")&"'")(0)
	if rs("userid3")<>"" and rs("userid3")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid3")&"'")(0)
	if left(str_sm,1)="/" then response.write mid(str_sm,2)
	%></td>
      <td align="center" bgcolor="#ffffff"><%
	hqallmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	if isnull(hqallmoney) then hqallmoney=0
	response.write Formatnumber(hqallmoney/count222,1,0,0,0)

    '���ڲ�����ֶ�
	customerlosttypeid = GetFieldDataBySQL("select customerlosttype from kehu where id="&rs("kehu_id"),"int",0)
	For aai = 1 To UBound(arr_cons_info, 1)
		If CInt(arr_cons_info(aai, 0))=customerlosttypeid Then 
		  arr2 = arr_cons_info(aai, 3)
		  arr3 = arr_cons_info(aai, 4)
		  If IsArray(arr2) Then 
			  for cci = 1 to ubound(arr2)
				if cint(arr2(cci))>=hqallmoney then
					arr3(cci) = arr3(cci) + hqallmoney/count222
					exit for
				end if
			  Next
			  arr_cons_info(aai, 3) = arr2
			  arr_cons_info(aai, 4) = arr3
			  Exit For 
		  End if
		End If 
	Next %></td>
      <td align="center"><%
	'money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and xiangmu_id in (select id from shejixiadan where ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"')")(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
	if rs("ky_name")<>cur_peplename then
			response.Write "/"&rs("ky_name")
	  end if
	  if rs("ky_name2")<>cur_peplename then
			response.Write "/"&rs("ky_name2")
	  end if
	%></td>
      <td align="center" bgcolor="#ffffff"><%if rs("ky_name")<>cur_peplename and rs("ky_name2")<>cur_peplename then
		response.write "0"
	else%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
          <tr>
            <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
            <td>&nbsp;<%=rsdg("all_sl")%></td>
          </tr>
          <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
        </table>
        <%end if%></td>
      <td align="center"><%=FinalMoneySum(rs("id"),False)%></td>
    </tr>
    <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
 
  rs.movenext
  i=i+1
loop
rs.close()
  %>
</table>
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;ѡƬ���� <%=Formatnumber(hq_indate_allsavemoney,1,0,0,0)%> Ԫ&nbsp;&nbsp;&nbsp; &nbsp;�ۼƺ���ѡƬǷ��
        <%
	tmp_fujia_money = conn.execute("select sum(f.money) from fujia f inner join shejixiadan s on f.xiangmu_id=s.id where (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"')")(0)
	tmp_save_money = conn.execute("select sum(m.money) from save_money m inner join shejixiadan s on m.xiangmu_id=s.id where m.type=2 and (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"')")(0)
	if isnull(tmp_fujia_money) then tmp_fujia_money = 0
	if isnull(tmp_save_money) then tmp_save_money = 0
	response.write Formatnumber(tmp_fujia_money-tmp_save_money,1,0,0,0)%>
        Ԫ&nbsp;&nbsp;&nbsp; &nbsp;������� <%=Formatnumber(ReceivablesMoney,1,0,0,0)%> Ԫ<br><%

	if IsArray(arr_cons_info) then
	response.write "&nbsp;���ڲ�����ֶκϼ�<br>"
	For aai = 1 To UBound(arr_cons_info, 1)
		arr1 = arr_cons_info(aai, 2)
		arr2 = arr_cons_info(aai, 3)
		arr3 = arr_cons_info(aai, 4)
		arr4 = arr_cons_info(aai, 5)
		If IsArray(arr1) Then 
			response.write "&nbsp;" & arr_cons_info(aai, 1) & "��"
			for cci = 1 to ubound(arr1)
				response.write arr1(cci)&" ~ "&arr2(cci)&"Ԫ("&arr4(cci)&")��"& Formatnumber(arr3(cci),1,0,0,0)&"Ԫ&nbsp;&nbsp;&nbsp;"
			Next
			response.write "<br>"
		End if
	Next 
	end if%></td>
    </tr>
  </table>
  <br>
  <%Call init_key()%>
<div align="center" style="line-height:30px">
  <%response.write datearea%>
  &nbsp; �����б�</div>
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" style="richness:1px">
    <tr bgcolor="#99FFFF">
      <td width="60" height="19" align="center">����</td>
      <td width="80" align="center">�ͻ�</td>
      <td align="center">��ϵ����</td>
      <td width="70" align="center">����ϵ��</td>
      <td width="200" align="center">��ϵ�ɷ�/(�Ŷ�)/
        ����/���ս�</td>
      <td width="120" align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font>/����</td>
      <td width="70" align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
      <td align="center">����ʱ��</td>
      <td width="70" align="center">����</td>
      <td align="center">��ע</td>
    </tr>
    <%
	rs.open "select * from shejixiadan where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&GetSqlCheckDateString("lc_cp"),conn,1,1
  do while not rs.eof
  str_sm=""
  	count111=1
	if not isnull(rs("userid2")) and rs("userid2")<>"" then count111=count111+1
  	if not isnull(rs("userid3")) and rs("userid3")<>"" then count111=count111+1
	
  MonthWedsuitCost = MonthWedsuitCost + getWedsuitCost(rs("id"))
	
  bk_jixiang=0
  bk_fujia=0
  
  '�������½ɺ��ڿ�
  hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
  	if isnull(money2) then money2=0
	count222 = 1
	if rs("ky_name2")<>"" and not isnull(rs("ky_name2")) then
		count222 = 2
	end if
	sm2_money=money2
	hq_indate_savemoney=hq_indate_savemoney/count222
  
  '�����ܺ���
  hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
  if isnull(hq_money) then hq_money = 0
  
  '�����ܺ��ڽɿ�
  hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
  
  'hq_minesavemoney = conn.execute("select sum(money) from save_money where [type]=2 and userid='"&userid&"' and xiangmu_id="&rs("id"))(0)
  
  set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
  do while not rshq.eof
  	if rshq("userid")=userid or rshq("userid2")=userid then
	  if rshq("userid")<>"" and not isnull(rshq("userid2")) then
		hq_mymoney = hq_mymoney + rshq("money")/2
	  else
	  	hq_mymoney = hq_mymoney + rshq("money")
  	  end if
	end if
	rshq.movenext
  loop
  rshq.close
  set rshq=nothing
  
  if isnull(hq_savemoney) then hq_savemoney = 0
  
  money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
  if isnull(money1) then money1=0
  
  '��Ƿ��
  hq_notsavemoney=hq_notsavemoney+hq_money-hq_savemoney
  
  '�ܺ���
  hq_allmoney=hq_allmoney+hq_money
  
  '�����ܺ��ڽɿ�
  hq_indate_allsavemoney=hq_indate_allsavemoney+hq_indate_savemoney
  customerlosttypeid = GetFieldDataBySQL("select customerlosttype from kehu where id="&rs("kehu_id"),"int",0)

'  For aai = 1 To UBound(arr_cons_info, 1)
'	If CInt(arr_cons_info(aai, 0))=customerlosttypeid Then 
'	  arr2 = arr_cons_info(aai, 3)
'	  arr3 = arr_cons_info(aai, 4)
'	  arr4 = arr_cons_info(aai, 5)
'	  If IsArray(arr2) Then 
'		  for cci = 1 to ubound(arr2)
'			if cint(arr2(cci))>=money1 then
'				arr3(cci) = arr3(cci) + money1/count111
'				arr4(cci) = arr4(cci) + 1
'				exit for
'			end if
'		  Next
'		  arr_cons_info(aai, 4) = arr3
'		  arr_cons_info(aai, 5) = arr4
'		  Exit For 
'	  End if
'	End If 
'  Next 
  %>
    <tr bgcolor="#FFFFFF" id="<%="tr_"&rs("id")%>">
      <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		msidlist=msidlist&rs("id")&","
	%>
      </td>
      <td align="center"><%
	 response.Write conn.execute("select lxpeple from kehu where id="&rs("kehu_id"))(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>
      <td align="center"><%=GetFieldDataBySQL("select jixiang from jixiang where id="&rs("jixiang"),"str","&nbsp;")%></td>
      <td align="center"><% 
		jx_money = rs("jixiang_money")
		AllXiangmuMoney = AllXiangmuMoney + jx_money
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
      <td align="center"><%
	sm1_money=money1/count111
	if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then response.Write formatnumber(sm1_money,1,0,0,0)
	if not isnull(rs("userid")) and rs("userid")<>"" and rs("userid")<>userid then str_sm=str_sm&"/"&GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid")&"'","str","N/A")
	if not isnull(rs("userid2")) and rs("userid2")<>"" and rs("userid2")<>userid then str_sm=str_sm&"/"&GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid2")&"'","str","N/A")
	if not isnull(rs("userid3")) and rs("userid3")<>"" and rs("userid3")<>userid then str_sm=str_sm&"/"&GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid3")&"'","str","N/A")
	if left(str_sm,1)="/" then response.write mid(str_sm,2)
	
	dd_dingjin=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and (beizhu='���𸶿�' or beizhu='���𸶿�')")(0)
	if isnull(dd_dingjin) then dd_dingjin=0
	dd_paizhao=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and beizhu<>'���𸶿�' and beizhu<>'���𸶿�'")(0)
	if isnull(dd_paizhao) then dd_paizhao=0
	response.write "/"&dd_dingjin&"/"&dd_paizhao
	%></td>
      <td align="center"><%
	money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and xiangmu_id in (select id from shejixiadan where ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"')")(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
	if rs("ky_name")<>cur_peplename then
			response.Write "/"&rs("ky_name")
	  end if
	  if rs("ky_name2")<>cur_peplename then
			response.Write "/"&rs("ky_name2")
	  end if
	%></td>
      <td align="center"><%qkmoney=FinalMoneySum(rs("id"),False)
	  AllQiankuanMoney = AllQiankuanMoney + qkmoney
	  response.write formatnumber(qkmoney,1,0,0,0)%></td>
      <td align="center">&nbsp;<%=datevalue(rs("times"))%></td>
      <td align="center"><%=getPerStep(rs("id"))%>&nbsp;</td>
      <td bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
  
  if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then
  	jixiang_money = jixiang_money + jx_money
  	money00=money00+money1
	tx_savemoney = conn.execute("select sum([money]) from save_money where [type]=1 and xiangmu_id="&rs("id"))(0)
  	if isnull(tx_savemoney) then tx_savemoney=0
  	if tx_savemoney=rs("jixiang_money") and conn.execute("select count(*) from save_money where xiangmu_id="&rs("id"))(0)>0 then
  		ReceivablesMoney = ReceivablesMoney + (rs("jixiang_money")/count111)
  	end if
  end if
  if hq_money=hq_indate_savemoney then 
  	RecFujiaMoney = RecFujiaMoney+hq_mymoney
	AllRecFujiaMoney = AllRecFujiaMoney+hq_money
  end if
  rs.movenext
  i=i+1
loop
rs.close
  %>
  </table>
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td> ��ϵ�ܽ�� <%=formatnumber(AllXiangmuMoney,1,0,0,0)%> Ԫ &nbsp;&nbsp;&nbsp;&nbsp;������Ƿ�� <%=formatnumber(AllQiankuanMoney,1,0,0,0)%>&nbsp;Ԫ </td>
    </tr>
</table>
<%
if instr(qj_flag,"1") then
  Call init_key()
	set rs=server.CreateObject("adodb.recordset")
	qj_sql="select * from shejixiadan where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and wc_name<>'' and not isnull(wc_name) and "&GetSqlCheckDateString("lc_wc")
	rs.open qj_sql,conn,1,1
%>
<div align="center" style="line-height:30px">
  <%response.write datearea%>
  &nbsp; ��ϵȡ���б�</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">����ϵ��</td>
    <td align="center">��ϵ�ɷ�/(�Ŷ�)</td>
    <td align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font>/����</td>
    <td align="center">ѡƬʱ��</td>
    <td align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
    <td align="center">��Ƭ���͡�</td>
    <td width="16%" align="center">��Ƭ���/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
  </tr>
  <%do while not rs.eof
	  str_sm=""
	  if not isnull(rs("userid3")) and rs("userid3")<>"" then 
		count111=3
		elseif not isnull(rs("userid2")) and rs("userid2")<>"" then
		count111=2
		else
		count111=1
		end if
	 
	  
	  '�������½ɺ��ڿ�
	  hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	  if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
		if isnull(money2) then money2=0
		count222 = 1
		if rs("ky_name2")<>"" and not isnull(rs("ky_name2")) then
			count222 = 2
		end if
		sm2_money=money2
		hq_indate_savemoney=hq_indate_savemoney/count222
	  
	  '�����ܺ���
	  hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	  if isnull(hq_money) then hq_money = 0
	  
	  '�����ܺ��ڽɿ�
	  hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
	  
	  'hq_minesavemoney = conn.execute("select sum(money) from save_money where [type]=2 and userid='"&userid&"' and xiangmu_id="&rs("id"))(0)
	  set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
	  do while not rshq.eof
		if rshq("userid")=userid or rshq("userid2")=userid then
		  if rshq("userid")<>"" and not isnull(rshq("userid2")) then
			hq_mymoney = hq_mymoney + rshq("money")/2
		  else
			hq_mymoney = hq_mymoney + rshq("money")
		  end if
		end if
		rshq.movenext
	  loop
	  rshq.close
	  set rshq=nothing
	  
	  if isnull(hq_savemoney) then hq_savemoney = 0
	  
	  '��Ƿ��
	  hq_notsavemoney=hq_notsavemoney+hq_money-hq_savemoney
	  
	  '�ܺ���
	  hq_allmoney=hq_allmoney+hq_money
	  
	  '�����ܺ��ڽɿ�
	  hq_indate_allsavemoney=hq_indate_allsavemoney+hq_indate_savemoney
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"	
	%>
    </td>
    <td align="center"><%
	 response.Write  conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>
    <td align="center"><% 
		jx_money = rs("jixiang_money")/count111
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
    <td align="center"><%money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money1) then money1=0
	if rs("userid")<>userid and rs("userid2")<>userid and rs("userid3")<>userid then money1=0
	sm1_money=money1/count111
	if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then response.Write formatnumber(sm1_money,1,0,0,0)
	if rs("userid")<>"" and rs("userid")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)
	if rs("userid2")<>"" and rs("userid2")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid2")&"'")(0)
	if rs("userid3")<>"" and rs("userid3")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid3")&"'")(0)
	if left(str_sm,1)="/" then response.write mid(str_sm,2)
	%>
    </td>
    <td align="center"><%
	money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and xiangmu_id in (select id from shejixiadan where ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"')")(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
	if rs("ky_name")<>cur_peplename then
			response.Write "/"&rs("ky_name")
	  end if
	  if rs("ky_name2")<>cur_peplename then
			response.Write "/"&rs("ky_name2")
	  end if
	%></td>
    <td align="center"><%if not isnull(rs("lc_ky")) then
		response.write datevalue(rs("lc_ky"))
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%=FinalMoneySum(rs("id"),False)%></td>
    <td align="center"><%if rs("ky_name")<>cur_peplename and rs("ky_name2")<>cur_peplename then
		response.write "0"
	else%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
          <tr>
            <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
            <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
          </tr>
          <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
        </table>
      <%end if%></td>
    <td align="center"><%
	dgallmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	if isnull(dgallmoney) then dgallmoney=0
	response.write formatnumber(dgallmoney/count222,1,0,0,0)
	%></td>
  </tr>
  <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
  
  if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then
  	jixiang_money = jixiang_money + jx_money
  	money00=money00+money1
	tx_savemoney = conn.execute("select sum([money]) from save_money where [type]=1 and xiangmu_id="&rs("id"))(0)
  	if isnull(tx_savemoney) then tx_savemoney=0
  	if tx_savemoney=rs("jixiang_money") and conn.execute("select count(*) from save_money where xiangmu_id="&rs("id"))(0)>0 then
  		ReceivablesMoney = ReceivablesMoney + (rs("jixiang_money")/count111)
  	end if
  end if
  if hq_money=hq_indate_savemoney then 
  	RecFujiaMoney = RecFujiaMoney+hq_mymoney
	AllRecFujiaMoney = AllRecFujiaMoney+hq_money
  end if
  rs.movenext
  i=i+1
loop
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td> &nbsp;��ϵ�ܽ�� <%=formatnumber(jixiang_money,1,0,0,0)%> Ԫ &nbsp;&nbsp;&nbsp;&nbsp;���½�����ϵ�� <%=formatnumber(ReceivablesMoney,1,0,0,0)%>&nbsp;Ԫ &nbsp;&nbsp;&nbsp;��ϵδ��
      <%
	response.write formatnumber(jixiang_money-money00,1,0,0,0)
'	jixiang_choucheng=money11*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
'	'response.write formatnumber(money11,1,0,0,0)
'	if isnull(jixiang_choucheng) then jixiang_choucheng=0
'	response.write formatnumber(jixiang_money-money11,1,0,0,0)%>
Ԫ&nbsp;<br>
&nbsp;�����ܽ�� <%=formatnumber(hq_mymoney,1,0,0,0)%> Ԫ&nbsp;&nbsp;&nbsp;&nbsp;(����)���½�����ڿ� <%=formatnumber(RecFujiaMoney,1,0,0,0)%>&nbsp;Ԫ &nbsp;&nbsp;&nbsp;(�Ŷ�)���½�����ڿ� <%=formatnumber(AllRecFujiaMoney,1,0,0,0)%>&nbsp;Ԫ &nbsp;&nbsp;&nbsp;����δ�� <%=formatnumber(hq_notsavemoney,1,0,0,0)%> Ԫ<br>
    &nbsp;</td>
  </tr>
</table>
<%
end if
if instr(qj_flag,"2") then
  Call init_key()
	set rs=server.CreateObject("adodb.recordset")
	qj_sql="select * from shejixiadan where (ky_name='"&peplename&"' or ky_name2='"&peplename&"') and wc_name<>'' and not isnull(wc_name) and "&GetSqlCheckDateString("lc_wc")
	'response.write qj_sql
	' and (userid<>'"&userid&"' and userid2<>'"&userid&"' and userid3<>'"&userid&"')
	rs.open qj_sql,conn,1,1
%>
<div align="center" style="line-height:30px">
  <%response.write datearea%>
  &nbsp; ����ȡ���б�</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">����ϵ��</td>
    <td align="center">��ϵ�ɷ�/(�Ŷ�)</td>
    <td align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font>/����</td>
    <td align="center">ѡƬʱ��</td>
    <td align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
    <td align="center">��Ƭ���͡�</td>
    <td width="16%" align="center">��Ƭ���/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
  </tr>
  <%do while not rs.eof
	  str_sm=""
	  if not isnull(rs("userid3")) and rs("userid3")<>"" then 
		count111=3
		elseif not isnull(rs("userid2")) and rs("userid2")<>"" then
		count111=2
		else
		count111=1
		end if
	 
	  
	  '�������½ɺ��ڿ�
	  hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	  if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
		if isnull(money2) then money2=0
		count222 = 1
		if rs("ky_name2")<>"" and not isnull(rs("ky_name2")) then
			count222 = 2
		end if
		sm2_money=money2
		hq_indate_savemoney=hq_indate_savemoney/count222
	  
	  '�����ܺ���
	  hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	  if isnull(hq_money) then hq_money = 0
	  
	  '�����ܺ��ڽɿ�
	  hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
	  
	  'hq_minesavemoney = conn.execute("select sum(money) from save_money where [type]=2 and userid='"&userid&"' and xiangmu_id="&rs("id"))(0)
	  set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
	  do while not rshq.eof
		if rshq("userid")=userid or rshq("userid2")=userid then
		  if rshq("userid")<>"" and not isnull(rshq("userid2")) then
			hq_mymoney = hq_mymoney + rshq("money")/2
		  else
			hq_mymoney = hq_mymoney + rshq("money")
		  end if
		end if
		rshq.movenext
	  loop
	  rshq.close
	  set rshq=nothing
	  
	  if isnull(hq_savemoney) then hq_savemoney = 0
	  
	  '��Ƿ��
	  hq_notsavemoney=hq_notsavemoney+hq_money-hq_savemoney
	  
	  '�ܺ���
	  hq_allmoney=hq_allmoney+hq_money
	  
	  '�����ܺ��ڽɿ�
	  hq_indate_allsavemoney=hq_indate_allsavemoney+hq_indate_savemoney
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"	
	%></td>
    <td align="center"><%
	 response.Write  conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>
    <td align="center"><% 
		jx_money = rs("jixiang_money")/count111
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
    <td align="center"><%money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money1) then money1=0
	if rs("userid")<>userid and rs("userid2")<>userid and rs("userid3")<>userid then money1=0
	sm1_money=money1/count111
	if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then response.Write formatnumber(sm1_money,1,0,0,0)
	if rs("userid")<>"" and rs("userid")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid")&"'")(0)
	if rs("userid2")<>"" and rs("userid2")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid2")&"'")(0)
	if rs("userid3")<>"" and rs("userid3")<>userid then str_sm=str_sm&"/"&conn.execute("select peplename from yuangong where username='"&rs("userid3")&"'")(0)
	if left(str_sm,1)="/" then response.write mid(str_sm,2)
	%></td>
    <td align="center"><%
	money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and xiangmu_id in (select id from shejixiadan where ky_name='"&cur_peplename&"' or ky_name2='"&cur_peplename&"')")(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
	if rs("ky_name")<>cur_peplename then
			response.Write "/"&rs("ky_name")
	  end if
	  if rs("ky_name2")<>cur_peplename then
			response.Write "/"&rs("ky_name2")
	  end if
	%></td>
    <td align="center"><%if not isnull(rs("lc_ky")) then
		response.write datevalue(rs("lc_ky"))
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%=FinalMoneySum(rs("id"),False)%></td>
    <td align="center"><%if rs("ky_name")<>cur_peplename and rs("ky_name2")<>cur_peplename then
		response.write "0"
	else%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
        <tr>
          <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
          <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
        </tr>
        <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
      </table>
      <%end if%></td>
    <td align="center"><%
	dgallmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	if isnull(dgallmoney) then dgallmoney=0
	response.write formatnumber(dgallmoney/count222,1,0,0,0)
	%></td>
  </tr>
  <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
  
  if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then
  	jixiang_money = jixiang_money + jx_money
  	money00=money00+money1
	tx_savemoney = conn.execute("select sum([money]) from save_money where [type]=1 and xiangmu_id="&rs("id"))(0)
  	if isnull(tx_savemoney) then tx_savemoney=0
  	if tx_savemoney=rs("jixiang_money") and conn.execute("select count(*) from save_money where xiangmu_id="&rs("id"))(0)>0 then
  		ReceivablesMoney = ReceivablesMoney + (rs("jixiang_money")/count111)
  	end if
  end if
  if hq_money=hq_indate_savemoney then 
  	RecFujiaMoney = RecFujiaMoney+hq_mymoney
	AllRecFujiaMoney = AllRecFujiaMoney+hq_money
  end if
  rs.movenext
  i=i+1
loop
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;��ϵ�ܽ�� <%=formatnumber(jixiang_money,1,0,0,0)%> Ԫ &nbsp;&nbsp;&nbsp;&nbsp;���½�����ϵ�� <%=formatnumber(ReceivablesMoney,1,0,0,0)%>&nbsp;Ԫ &nbsp;&nbsp;&nbsp;��ϵδ��
      <%
	response.write formatnumber(jixiang_money-money00,1,0,0,0)
'	jixiang_choucheng=money11*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
'	'response.write formatnumber(money11,1,0,0,0)
'	if isnull(jixiang_choucheng) then jixiang_choucheng=0
'	response.write formatnumber(jixiang_money-money11,1,0,0,0)%>
      Ԫ&nbsp;<br>
      &nbsp;�����ܽ�� <%=formatnumber(hq_mymoney,1,0,0,0)%> Ԫ&nbsp;&nbsp;&nbsp;&nbsp;(����)���½�����ڿ� <%=formatnumber(RecFujiaMoney,1,0,0,0)%>&nbsp;Ԫ &nbsp;&nbsp;&nbsp;(�Ŷ�)���½�����ڿ� <%=formatnumber(AllRecFujiaMoney,1,0,0,0)%>&nbsp;Ԫ &nbsp;&nbsp;&nbsp;����δ�� <%=formatnumber(hq_notsavemoney,1,0,0,0)%> Ԫ<br>
      &nbsp;</td>
  </tr>
</table>
<%end if%>
<%
Call showSubTable()
case 2
dim dict_xc
set dict_xc_name=Server.CreateObject("Scripting.Dictionary")
set dict_xc_vol=Server.CreateObject("Scripting.Dictionary")
set dict_fd_name=Server.CreateObject("Scripting.Dictionary")
set dict_fd_vol=Server.CreateObject("Scripting.Dictionary")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where sj_name='"&peplename&"' and "&GetSqlCheckDateString("lc_sj"),conn,1,1
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19"><div align="left">&nbsp;&nbsp;����</div></td>
    <td><div align="center">�ͻ�/���� </div></td>
    <td><div align="center">��ϵ���</div></td>
   
	<td align="center">ѡƬ���</td>
	<td><div align="center">����</div></td>
	<td><div align="center">�Ŵ�</div></td>
   
    <td align="center" valign="middle">������</td>
    <td align="center" valign="middle">���淽ʽ</td>
  </tr>
  <%
   banmianll=0
   fangdall=0
   allxpnum=0
   xpcount=rs.recordcount
  do while not rs.eof
  	allxpnum = allxpnum + rs("sl2")
 	save_money=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&"")(0)
	
	if isnull(save_money) then save_money=0
	fujia1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia1) then fujia1=0
	fujia2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia2) then fujia2=0
	goumai=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
	if isnull(goumai) then goumai=0
	jx_money=rs("jixiang_money")
	money111=fujia1+fujia2+jx_money-save_money
	%>
  <tr bgcolor="#FFFFFF">
    <td>
      <div align="left"> &nbsp;
          <% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
	if idlist="" or isnull(idlist) then
		idlist=rs("id")
	else
		idlist=idlist&", "&rs("id")
	end if
	%>
    </div></td>
    <td><div align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%>/<%if money111>0 then 
	response.Write "δ����"
	else
	response.Write "�ѽ���"
	end if
	%>
	</div></td>
    <td><div align="center"><%
	response.write rs("jixiang_money")
	%></div></td>
   
    <td align="center"><%
  	hq_fujia=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id"))(0)
	  if isnull(hq_fujia) then hq_fujia=0
	  allhqmoney=allhqmoney+hq_fujia
	  response.Write cint(hq_fujia)&"Ԫ"%></td>
		<td align="center"><table width="85%"  border="0" cellspacing="0" cellpadding="0">
       <%
	if not isnull(rs("yunyong")) and rs("yunyong")<>"" then
		arr_idlist=split(rs("yunyong"),", ")
	  dim count11,count22,rslistflag
	  count11=ubound(arr_idlist)+1
	  if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			sllist=split(rs("pagevol"),", ")
	  end if
	  count22=0
	  for yy=1 to count11
		
		set rslistflag = conn.execute("select [isxc] from yunyong where id="&arr_idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("isxc")=1 then
				count22=count22+1
	%>
        <tr><td><%
		dim yyflag,rslist_yunyong
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&arr_idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			tmp_xcvol = cint(sllist(yy-1))
		else
			tmp_xcvol = 0
		end if
		
		if dict_xc_name(arr_idlist(yy-1))<>"" then
			dict_xc_vol(arr_idlist(yy-1))=dict_xc_vol(arr_idlist(yy-1))+tmp_xcvol
		else
			dict_xc_name(arr_idlist(yy-1))=rslist_yunyong("yunyong")
			dict_xc_vol(arr_idlist(yy-1))=tmp_xcvol
		end if
		response.write tmp_xcvol
		response.write "P"
		rslist_yunyong.close()
		%></td> </tr>
        <%
			end if
			end if
			rslistflag.close()
		next
	end if
	
	dim rslist_fujia
	set rslist_fujia = conn.execute("select fujia.jixiang,fujia.pagevol from fujia inner join yunyong on fujia.jixiang=yunyong.id where fujia.xiangmu_id="&rs("id")&" and yunyong.isxc=1")
	do while not rslist_fujia.eof
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&rslist_fujia("jixiang"))
		response.Write "<tr><td>"&rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_xc_name(cstr(rslist_fujia("jixiang")))<>"" then
			dict_xc_vol(cstr(rslist_fujia("jixiang")))=dict_xc_vol(cstr(rslist_fujia("jixiang")))+cint(rslist_fujia("pagevol"))
		else
			dict_xc_name(cstr(rslist_fujia("jixiang")))=rslist_yunyong("yunyong")
			dict_xc_vol(cstr(rslist_fujia("jixiang")))=cint(rslist_fujia("pagevol"))
		end if
		
		response.write rslist_fujia("pagevol")
		response.write "P"
		response.write "</td></tr>"
		rslist_yunyong.close()
		rslist_fujia.movenext
	loop
	rslist_fujia.close
	set rslist_fujia = nothing
		%>
      
    </table></td>
		<td align="center"><table width="85%"  border="0" cellspacing="0" cellpadding="0">
          <%
	if not isnull(rs("yunyong")) and rs("yunyong")<>"" then
		arr_idlist=split(rs("yunyong"),", ")
		arr_sllist=split(rs("sl"),", ")
	  count11=ubound(arr_idlist)+1
	  count22=0
	  for yy=1 to count11
		
		set rslistflag = conn.execute("select [type4] from yunyong where id="&arr_idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("type4")=1 then
				count22=count22+1
	%>
          <tr>
            <td><%
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&arr_idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_fd_name(arr_idlist(yy-1))<>"" then
			dict_fd_vol(arr_idlist(yy-1))=dict_fd_vol(arr_idlist(yy-1))+cint(arr_sllist(yy-1))
		else
			dict_fd_name(arr_idlist(yy-1))=rslist_yunyong("yunyong")
			dict_fd_vol(arr_idlist(yy-1))=cint(arr_sllist(yy-1))
		end if

		response.write arr_sllist(yy-1)
		response.write "��"
		rslist_yunyong.close()
		%></td>
          </tr>
          <%
			end if
			end if
			rslistflag.close()
		next
	end if
	
	set rslist_fujia = conn.execute("select fujia.jixiang,fujia.sl from fujia inner join yunyong on fujia.jixiang=yunyong.id where fujia.xiangmu_id="&rs("id")&" and yunyong.type4=1")
	do while not rslist_fujia.eof
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&rslist_fujia("jixiang"))
		response.Write "<tr><td>"&rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_fd_name(cstr(rslist_fujia("jixiang")))<>"" then
			dict_fd_vol(cstr(rslist_fujia("jixiang")))=dict_fd_vol(cstr(rslist_fujia("jixiang")))+cint(rslist_fujia("sl"))
		else
			dict_fd_name(cstr(rslist_fujia("jixiang")))=rslist_yunyong("yunyong")
			dict_fd_vol(cstr(rslist_fujia("jixiang")))=cint(rslist_fujia("sl"))
		end if
		
		response.write rslist_fujia("sl")
		response.write "��"
		response.write "</td></tr>"
		rslist_yunyong.close()
		rslist_fujia.movenext
	loop
	rslist_fujia.close
	set rslist_fujia = nothing
	%>
    </table></td>
		<td align="center"><%if not isnull(rs("lc_sj")) then
		response.write datevalue(rs("lc_sj"))
	else
		response.write "&nbsp;"
	end if%></td>
        <td align="center"><%
		if rs("xg_opt")=0 then
			response.write "�ڲ�����"
		else
			response.Write "���˿���"
		end if
		%></td>
  </tr>
  <%
 ' choucheng11=choucheng11+choucheng
   banmianll=banmianll+banmian
  fangdall=fangdall+fangda
 
  jixiang_money=jixiang_money+rs("jixiang_money")
  rs.movenext
  i=i+1
loop
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;����ϵ <%=formatnumber(jixiang_money,1,0,0,0)%>Ԫ&nbsp;&nbsp; �ܺ��� <%=formatnumber(allhqmoney,1,0,0,0)%>Ԫ&nbsp;&nbsp; �������: <%=num13%>��&nbsp;&nbsp;&nbsp; ������ϵ��Ƭ������<%=allxpnum%> ��&nbsp;&nbsp;&nbsp; ������˴�����<%=xpcount%> ��<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="100" valign="top">&nbsp;��Ƭ��Ŀ�б�<br>&nbsp;���ܼ�<span id="sp_gp">0</span>����</td>
    <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
	  if idlist="" or isnull(idlist) then
	  	response.write "<td>��</td>"
	  else
	  set rs_dg=server.createobject("adodb.recordset")
	  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&idlist&") and jixiang in (select id from yunyong where isgp=1) group by jixiang"
	  rs_dg.open sql,conn,1,1
	  dim gpvol
	  gpvol=0
	  if not rs_dg.eof then
	  For i=1 to rs_dg.recordcount 
	  If rs_dg.eof Then Exit For
	  gpvol=gpvol+rs_dg("all_sl")
	  %>
        <td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%>��</td>
        <%
		if i mod 5=0 then
			response.write "</tr><tr>"
		end if
		rs_dg.Movenext
	next
	end if
	response.write "<script language='javascript'>document.getElementById('sp_gp').innerHTML='"&gpvol&"';</script>"
	rs_dg.close
	set rs_dg=nothing
	end if
    %>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="100" valign="top">&nbsp;�����Ŀ�б�<br>&nbsp;���ܼ�<span id="sp_xc">0</span>P��</td>
    <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
		i=0
	  dim xcvol
	  xcvol=0
	  if dict_xc_name.Count>0 then
	  	for each idno in dict_xc_name
	  %>
        <td><%=dict_xc_name(idno)%>:&nbsp;<%=dict_xc_vol(idno)%>P</td>
        <%
			i=i+1
			xcvol=xcvol+cint(dict_xc_vol(idno))
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
	  else
	    response.write "<td>��</td>"
      end if
	  response.write "<script language='javascript'>document.getElementById('sp_xc').innerHTML='"&xcvol&"';</script>"
	set dict_xc_name=nothing
	set dict_xc_vol=nothing
    %>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="100" valign="top">&nbsp;�Ŵ���Ŀ�б�<br>&nbsp;���ܼ�<span id="sp_fd">0</span>�ţ�</td>
    <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
		i=0
	  dim fdvol
	  fdvol=0
	  if dict_fd_name.Count>0 then
	  	for each idno in dict_fd_name
	  %>
        <td><%=dict_fd_name(idno)%>:&nbsp;<%=dict_fd_vol(idno)%>��</td>
        <%
			i=i+1
			fdvol=fdvol+cint(dict_fd_vol(idno))
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
	  else
	    response.write "<td>��</td>"
      end if
	  response.write "<script language='javascript'>document.getElementById('sp_fd').innerHTML='"&fdvol&"';</script>"
	set dict_fd_name=nothing
	set dict_fd_vol=nothing
    %>
      </tr>
    </table></td>
  </tr>
</table></td>
  </tr>
</table>
<br>
<%
jixiang_money=0
set rs6=server.CreateObject("adodb.recordset")
rs6.open "select * from shejixiadan where xp_name='"&cur_peplename&"' and "&GetSqlCheckDateString("lc_ky"),conn,1,1
xpcount = rs6.recordcount
%>
<div align="center" style="line-height:30px">
  <%response.write datearea%>
  &nbsp;
��ɫ����</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td width="15%" height="19">&nbsp;&nbsp;����</td>
    <td width="12%" align="center">��ϵ�۸�</td>
    <td width="18%" align="center">�ͻ�</td>
    <td align="center">������Ŀ</td>
    <td width="12%" align="center">��Ƭ���/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td width="10%" align="center">��ɫ����</td>
    <td width="10%" align="center">��ϵ����</td>
  </tr>
  <%
  allxpnum=0
  alltsnum=0
  idlist=""
  do while not rs6.eof
  		alltsnum = alltsnum + rs6("tsVolume")
		allxpnum = allxpnum + rs6("sl2")
		jixiang_money = jixiang_money + rs6("jixiang_money")
  %>
  <tr bgcolor="#FFFFFF">
    <td> &nbsp;
          <% response.Write rs6("id")
	if idlist="" or isnull(idlist) then
		idlist=rs6("id")
	else
		idlist=idlist&", "&rs6("id")
	end if
	%>    </td>
    <td align="center"><%=rs6("jixiang_money")%></td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs6("kehu_id")&"")(0)%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs6("id")&" and "&GetSqlCheckDateString("times")&" group by jixiang")
	do while not rsdg.eof
	%>
      <tr>
        <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
        <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
        <td>&nbsp;<%=rsdg("all_money")%>Ԫ&nbsp;</td>
      </tr>
      <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
    </table></td>
    <td align="center">
      <%
	  dgmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs6("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	  if isnull(dgmoney) then dgmoney=0
	  response.write dgmoney
	money13=conn.execute("select sum(dj*sl) from sell_jilu where "&GetSqlCheckDateString("times")&"")(0)
	if isnull(money13) then money13=0
	money13=formatnumber(money13,1,0,0,0)
	%>    </td>
    <td align="center"><%=rs6("tsVolume")%></td>
    <td align="center"><%=rs6("sl2")%></td>
  </tr>
  <%
	money113=money113+money13
	'fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=2 and xiangmu_id="&rs6("id")&"")(0)
	'if isnull(fujia_save) then fujia_save=0
	'fujia_save11=fujia_save11+fujia_save
  rs6.movenext
  i=i+1
loop

  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;����ϵ��<%=jixiang_money%> Ԫ&nbsp;&nbsp;���º����տ
      <%'response.Write formatnumber(allsavemoney,1,0,0,0)
	  fujia_save11 = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and j.xp_name='"&cur_peplename&"' and "&GetSqlCheckDateString("s.times")&" and "&GetSqlCheckDateString("j.lc_ky"))(0)
	  if isnull(fujia_save11) then fujia_save11=0
	  hqbk_money = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and j.xp_name='"&cur_peplename&"' and "&GetSqlCheckDateString("s.times")&" and s.xiangmu_id not in (select id from shejixiadan where "&GetSqlCheckDateString("lc_ky")&")")(0)
	  if isnull(hqbk_money) then hqbk_money  = 0
	  response.Write formatnumber(fujia_save11,1,0,0,0)&" + "& hqbk_money &" (���ڲ���)"%>
Ԫ&nbsp; &nbsp;������ϵ��Ƭ������<%=allxpnum%> ��&nbsp;&nbsp;���µ�ɫ��ϵ������<%=alltsnum%> ��&nbsp;&nbsp;&nbsp; ������˴�����<%=xpcount%> ��</td>
  </tr>
</table>
<%call ShowSuitType(idlist)%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="15%" valign="top">&nbsp;��Ƭ��Ŀ�б�</td>
    <td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
	  if idlist="" or isnull(idlist) then
	  	response.write "<td>��</td>"
	  else
		  set rs_dg=server.createobject("adodb.recordset")
		  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&idlist&") and jixiang in (select id from yunyong where isgp=1) group by jixiang"
		  rs_dg.open sql,conn,1,1
		  if not rs_dg.eof then
		  For i=1 to rs_dg.recordcount 
		  If rs_dg.eof Then Exit For
		  %>
			<td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%> ��</td>
			<%
		if i mod 5=0 then
		response.write "</tr><tr>"
		end if
		rs_dg.Movenext
		next
		end if
		rs_dg.close
		set rs_dg=nothing
    end if%>
      </tr>
    </table></td>
  </tr>
</table>
<%
jixiang_money=0
set rs6=server.CreateObject("adodb.recordset")
rs6.open "select * from shejixiadan where jx_name='"&cur_peplename&"' and "&GetSqlCheckDateString("lc_jx"),conn,1,1
xpcount = rs6.recordcount
%>
<div align="center" style="line-height:30px">
  <%response.write datearea%>
  &nbsp;
  ���ޱ���</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td width="15%" height="19">&nbsp;&nbsp;����</td>
    <td width="12%" align="center">��ϵ�۸�</td>
    <td width="18%" align="center">�ͻ�</td>
    <td align="center">������Ŀ</td>
    <td width="12%" align="center">��Ƭ���/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td width="10%" align="center">��ɫ����</td>
    <td width="10%" align="center">��ϵ����</td>
  </tr>
  <%
  allxpnum=0
  alltsnum=0
  idlist=""
  do while not rs6.eof
  		alltsnum = alltsnum + rs6("tsVolume")
		allxpnum = allxpnum + rs6("sl2")
		jixiang_money = jixiang_money + rs6("jixiang_money")
  %>
  <tr bgcolor="#FFFFFF">
    <td>&nbsp;
        <% response.Write rs6("id")
	if idlist="" or isnull(idlist) then
		idlist=rs6("id")
	else
		idlist=idlist&", "&rs6("id")
	end if
	%>
    </td>
    <td align="center"><%=rs6("jixiang_money")%></td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs6("kehu_id")&"")(0)%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs6("id")&" and "&GetSqlCheckDateString("times")&" group by jixiang")
	do while not rsdg.eof
	%>
      <tr>
        <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
        <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
        <td>&nbsp;<%=rsdg("all_money")%>Ԫ&nbsp;</td>
      </tr>
      <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
    </table></td>
    <td align="center"><%
	  dgmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs6("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	  if isnull(dgmoney) then dgmoney=0
	  response.write dgmoney
	money13=conn.execute("select sum(dj*sl) from sell_jilu where "&GetSqlCheckDateString("times")&"")(0)
	if isnull(money13) then money13=0
	money13=formatnumber(money13,1,0,0,0)
	%>
    </td>
    <td align="center"><%=rs6("tsVolume")%></td>
    <td align="center"><%=rs6("sl2")%></td>
  </tr>
  <%
	money113=money113+money13
	'fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=2 and xiangmu_id="&rs6("id")&"")(0)
	'if isnull(fujia_save) then fujia_save=0
	'fujia_save11=fujia_save11+fujia_save
  rs6.movenext
  i=i+1
loop

  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;����ϵ��<%=jixiang_money%> Ԫ&nbsp;&nbsp;���º����տ
      <%'response.Write formatnumber(allsavemoney,1,0,0,0)
	  fujia_save11 = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and j.jx_name='"&cur_peplename&"' and "&GetSqlCheckDateString("s.times")&" and "&GetSqlCheckDateString("j.lc_ky"))(0)
	  if isnull(fujia_save11) then fujia_save11=0
	  hqbk_money = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and j.jx_name='"&cur_peplename&"' and "&GetSqlCheckDateString("s.times")&" and s.xiangmu_id not in (select id from shejixiadan where "&GetSqlCheckDateString("lc_ky")&")")(0)
	  if isnull(hqbk_money) then hqbk_money  = 0
	  response.Write formatnumber(fujia_save11,1,0,0,0)&" + "& hqbk_money &" (���ڲ���)"%>
      Ԫ&nbsp;</td>
  </tr>
</table>
<%
if instr(qj_flag,"1") then
  Call init_key()
	set rs=server.CreateObject("adodb.recordset")
	qj_sql="select * from shejixiadan where sj_name='"&peplename&"' and wc_name<>'' and not isnull(wc_name) and "&GetSqlCheckDateString("lc_wc")
	rs.open qj_sql,conn,1,1
	xpcount=rs.recordcount
	
	set dict_xc_name=Server.CreateObject("Scripting.Dictionary")
	set dict_xc_vol=Server.CreateObject("Scripting.Dictionary")
	set dict_fd_name=Server.CreateObject("Scripting.Dictionary")
	set dict_fd_vol=Server.CreateObject("Scripting.Dictionary")
%>

<div align="center" style="line-height:30px">
  <%response.write datearea%>
&nbsp; ���ȡ���б�</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF" align="center">
    <td width="7%" height="19">����</td>
    <td width="12%">�ͻ�/���� </td>
    <td>��ϵ</td>
    <td>����</td>
    <td>�Ŵ�</td>
    <td width="14%" align="center" valign="middle">���ȡ��</td>
  </tr>
  <%
   banmianll=0
   fangdall=0
   allxpnum =0
  do while not rs.eof
 	save_money=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&"")(0)
	if isnull(save_money) then save_money=0
	fujia1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia1) then fujia1=0
	fujia2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia2) then fujia2=0
	goumai=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
	if isnull(goumai) then goumai=0
	money111=fujia1+fujia2+rs("jixiang_money")-save_money
	allhqmoney=allhqmoney+fujia1
	
	banmian=rs("banmian")
	if isnull(banmian) then banmian=0
	 fangda=rs("fangda")
	if isnull(fangda) then fangda=0
	allxpnum = allxpnum + rs("sl2")
	 %>
  <tr bgcolor="#FFFFFF" align="center">
    <td><% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"%></td>
    <td><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%>/
    <%if money111>0 then 
	response.Write "δ����"
	else
	response.Write "�ѽ���"
	end if
	%></td>
    <td><%=GetFieldDataBySQL("select jixiang from jixiang where id="&rs("jixiang")&"","str","&nbsp;")%></td>

    <td><table width="95%"  border="0" cellspacing="0" cellpadding="0">
      <%
	if not isnull(rs("yunyong")) and rs("yunyong")<>"" then
		arr_idlist=split(rs("yunyong"),", ")
	  count11=ubound(arr_idlist)+1
	  if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			sllist=split(rs("pagevol"),", ")
	  end if
	  count22=0
	  for yy=1 to count11
		
		set rslistflag = conn.execute("select [isxc] from yunyong where id="&arr_idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("isxc")=1 then
				count22=count22+1
	%>
      <tr>
        <td><%
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&arr_idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			tmp_xcvol = cint(sllist(yy-1))
		else
			tmp_xcvol = 0
		end if
		
		if dict_xc_name(arr_idlist(yy-1))<>"" then
			dict_xc_vol(arr_idlist(yy-1))=dict_xc_vol(arr_idlist(yy-1))+tmp_xcvol
		else
			dict_xc_name(arr_idlist(yy-1))=rslist_yunyong("yunyong")
			dict_xc_vol(arr_idlist(yy-1))=tmp_xcvol
		end if
		response.write tmp_xcvol
		response.write "P"
		rslist_yunyong.close()
		%></td>
      </tr>
      <%
			end if
			end if
			rslistflag.close()
		next
	end if
	
	
	set rslist_fujia = conn.execute("select fujia.jixiang,fujia.pagevol from fujia inner join yunyong on fujia.jixiang=yunyong.id where fujia.xiangmu_id="&rs("id")&" and yunyong.isxc=1")
	do while not rslist_fujia.eof
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&rslist_fujia("jixiang"))
		response.Write "<tr><td>"&rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_xc_name(cstr(rslist_fujia("jixiang")))<>"" then
			dict_xc_vol(cstr(rslist_fujia("jixiang")))=dict_xc_vol(cstr(rslist_fujia("jixiang")))+cint(rslist_fujia("pagevol"))
		else
			dict_xc_name(cstr(rslist_fujia("jixiang")))=rslist_yunyong("yunyong")
			dict_xc_vol(cstr(rslist_fujia("jixiang")))=cint(rslist_fujia("pagevol"))
		end if
		
		response.write rslist_fujia("pagevol")
		response.write "P"
		response.write "</td></tr>"
		rslist_yunyong.close()
		rslist_fujia.movenext
	loop
	rslist_fujia.close
	set rslist_fujia = nothing
		%>
    </table></td>
    <td><table width="95%"  border="0" cellspacing="0" cellpadding="0">
      <%
	if not isnull(rs("yunyong")) and rs("yunyong")<>"" then
		arr_idlist=split(rs("yunyong"),", ")
		arr_sllist=split(rs("sl"),", ")
	  count11=ubound(arr_idlist)+1
	  count22=0
	  for yy=1 to count11
		
		set rslistflag = conn.execute("select [type4] from yunyong where id="&arr_idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("type4")=1 then
				count22=count22+1
	%>
      <tr>
        <td><%
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&arr_idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_fd_name(arr_idlist(yy-1))<>"" then
			dict_fd_vol(arr_idlist(yy-1))=dict_fd_vol(arr_idlist(yy-1))+cint(arr_sllist(yy-1))
		else
			dict_fd_name(arr_idlist(yy-1))=rslist_yunyong("yunyong")
			dict_fd_vol(arr_idlist(yy-1))=cint(arr_sllist(yy-1))
		end if

		response.write arr_sllist(yy-1)
		response.write "��"
		rslist_yunyong.close()
		%></td>
      </tr>
      <%
			end if
			end if
			rslistflag.close()
		next
	end if
	
	set rslist_fujia = conn.execute("select fujia.jixiang,fujia.sl from fujia inner join yunyong on fujia.jixiang=yunyong.id where fujia.xiangmu_id="&rs("id")&" and yunyong.type4=1")
	do while not rslist_fujia.eof
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&rslist_fujia("jixiang"))
		response.Write "<tr><td>"&rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_fd_name(cstr(rslist_fujia("jixiang")))<>"" then
			dict_fd_vol(cstr(rslist_fujia("jixiang")))=dict_fd_vol(cstr(rslist_fujia("jixiang")))+cint(rslist_fujia("sl"))
		else
			dict_fd_name(cstr(rslist_fujia("jixiang")))=rslist_yunyong("yunyong")
			dict_fd_vol(cstr(rslist_fujia("jixiang")))=cint(rslist_fujia("sl"))
		end if
		
		response.write rslist_fujia("sl")
		response.write "��"
		response.write "</td></tr>"
		rslist_yunyong.close()
		rslist_fujia.movenext
	loop
	rslist_fujia.close
	set rslist_fujia = nothing
	%>
    </table></td>
    <td width="14%"><%if not isnull(rs("lc_wc")) then
		response.write datevalue(rs("lc_wc"))
	else
		response.write "&nbsp;"
	end if%></td>
  </tr>
  <%
 ' choucheng11=choucheng11+choucheng
   banmianll=banmianll+banmian
  fangdall=fangdall+fangda
 
  jixiang_money=jixiang_money+rs("jixiang_money")
  rs.movenext
  i=i+1
loop
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;����ϵ <%=formatnumber(jixiang_money,1,0,0,0)%>Ԫ&nbsp;&nbsp; �ܺ��� <%=formatnumber(allhqmoney,1,0,0,0)%>Ԫ&nbsp;&nbsp; ������ϵ��Ƭ������<%=allxpnum%> ��&nbsp;&nbsp;&nbsp; ������˴�����<%=xpcount%> ��<br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#FFFFFF">
          <td width="100" valign="top">&nbsp;��Ƭ��Ŀ�б�<br>
            &nbsp;���ܼ�<span id="sp_gp2">0</span>����</td>
          <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <%
	  if idlist="" or isnull(idlist) then
	  	response.write "<td>��</td>"
	  else
	  set rs_dg=server.createobject("adodb.recordset")
	  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&idlist&") and jixiang in (select id from yunyong where isgp=1) group by jixiang"
	  rs_dg.open sql,conn,1,1
	  gpvol=0
	  if not rs_dg.eof then
	  For i=1 to rs_dg.recordcount 
	  If rs_dg.eof Then Exit For
	  gpvol=gpvol+rs_dg("all_sl")
	  %>
              <td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%>��</td>
              <%
		if i mod 5=0 then
			response.write "</tr><tr>"
		end if
		rs_dg.Movenext
	next
	end if
	response.write "<script language='javascript'>document.getElementById('sp_gp2').innerHTML='"&gpvol&"';</script>"
	rs_dg.close
	set rs_dg=nothing
	end if
    %>
            </tr>
          </table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#FFFFFF">
          <td width="100" valign="top">&nbsp;�����Ŀ�б�<br>
            &nbsp;���ܼ�<span id="sp_xc2">0</span>P��</td>
          <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <%
		i=0
	  xcvol=0
	  if dict_xc_name.Count>0 then
	  	for each idno in dict_xc_name
	  %>
              <td><%=dict_xc_name(idno)%>:&nbsp;<%=dict_xc_vol(idno)%>P</td>
              <%
			i=i+1
			xcvol=xcvol+cint(dict_xc_vol(idno))
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
	  else
	    response.write "<td>��</td>"
      end if
	  response.write "<script language='javascript'>document.getElementById('sp_xc2').innerHTML='"&xcvol&"';</script>"
	set dict_xc_name=nothing
	set dict_xc_vol=nothing
    %>
            </tr>
          </table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#FFFFFF">
          <td width="100" valign="top">&nbsp;�Ŵ���Ŀ�б�<br>
            &nbsp;���ܼ�<span id="sp_fd2">0</span>�ţ�</td>
          <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <%
		i=0
	  fdvol=0
	  if dict_fd_name.Count>0 then
	  	for each idno in dict_fd_name
	  %>
              <td><%=dict_fd_name(idno)%>:&nbsp;<%=dict_fd_vol(idno)%>��</td>
              <%
			i=i+1
			fdvol=fdvol+cint(dict_fd_vol(idno))
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
	  else
	    response.write "<td>��</td>"
      end if
	  response.write "<script language='javascript'>document.getElementById('sp_fd2').innerHTML='"&fdvol&"';</script>"
	set dict_fd_name=nothing
	set dict_fd_vol=nothing
    %>
            </tr>
          </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<%end if%>
<%
if instr(qj_flag,"2") then
  Call init_key()
	set rs=server.CreateObject("adodb.recordset")
	qj_sql="select * from shejixiadan where xp_name='"&peplename&"' and wc_name<>'' and not isnull(wc_name) and "&GetSqlCheckDateString("lc_wc")
	rs.open qj_sql,conn,1,1
	xpcount=rs.recordcount
	
	set dict_xc_name=Server.CreateObject("Scripting.Dictionary")
	set dict_xc_vol=Server.CreateObject("Scripting.Dictionary")
	set dict_fd_name=Server.CreateObject("Scripting.Dictionary")
	set dict_fd_vol=Server.CreateObject("Scripting.Dictionary")
%>

<div align="center" style="line-height:30px">
  <%response.write datearea%>
&nbsp; ��ɫȡ���б�</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF" align="center">
    <td width="7%" height="19">����</td>
    <td width="12%">�ͻ�/���� </td>
    <td>��ϵ</td>
    <td>����</td>
    <td>�Ŵ�</td>
    <td width="14%" align="center" valign="middle">���ȡ��</td>
  </tr>
  <%
   banmianll=0
   fangdall=0
   allxpnum =0
  do while not rs.eof
 	save_money=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&"")(0)
	if isnull(save_money) then save_money=0
	fujia1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia1) then fujia1=0
	fujia2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia2) then fujia2=0
	goumai=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
	if isnull(goumai) then goumai=0
	money111=fujia1+fujia2+rs("jixiang_money")-save_money
	allhqmoney=allhqmoney+fujia1
	
	banmian=rs("banmian")
	if isnull(banmian) then banmian=0
	 fangda=rs("fangda")
	if isnull(fangda) then fangda=0
	allxpnum = allxpnum + rs("sl2")
	 %>
  <tr bgcolor="#FFFFFF" align="center">
    <td><% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"%></td>
    <td><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%>/
    <%if money111>0 then 
	response.Write "δ����"
	else
	response.Write "�ѽ���"
	end if
	%></td>
    <td><%=GetFieldDataBySQL("select jixiang from jixiang where id="&rs("jixiang")&"","str","&nbsp;")%></td>

    <td><table width="95%"  border="0" cellspacing="0" cellpadding="0">
      <%
	if not isnull(rs("yunyong")) and rs("yunyong")<>"" then
		arr_idlist=split(rs("yunyong"),", ")
	  count11=ubound(arr_idlist)+1
	  if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			sllist=split(rs("pagevol"),", ")
	  end if
	  count22=0
	  for yy=1 to count11
		
		set rslistflag = conn.execute("select [isxc] from yunyong where id="&arr_idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("isxc")=1 then
				count22=count22+1
	%>
      <tr>
        <td><%
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&arr_idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			tmp_xcvol = cint(sllist(yy-1))
		else
			tmp_xcvol = 0
		end if
		
		if dict_xc_name(arr_idlist(yy-1))<>"" then
			dict_xc_vol(arr_idlist(yy-1))=dict_xc_vol(arr_idlist(yy-1))+tmp_xcvol
		else
			dict_xc_name(arr_idlist(yy-1))=rslist_yunyong("yunyong")
			dict_xc_vol(arr_idlist(yy-1))=tmp_xcvol
		end if
		response.write tmp_xcvol
		response.write "P"
		rslist_yunyong.close()
		%></td>
      </tr>
      <%
			end if
			end if
			rslistflag.close()
		next
	end if
	
	
	set rslist_fujia = conn.execute("select fujia.jixiang,fujia.pagevol from fujia inner join yunyong on fujia.jixiang=yunyong.id where fujia.xiangmu_id="&rs("id")&" and yunyong.isxc=1")
	do while not rslist_fujia.eof
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&rslist_fujia("jixiang"))
		response.Write "<tr><td>"&rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_xc_name(cstr(rslist_fujia("jixiang")))<>"" then
			dict_xc_vol(cstr(rslist_fujia("jixiang")))=dict_xc_vol(cstr(rslist_fujia("jixiang")))+cint(rslist_fujia("pagevol"))
		else
			dict_xc_name(cstr(rslist_fujia("jixiang")))=rslist_yunyong("yunyong")
			dict_xc_vol(cstr(rslist_fujia("jixiang")))=cint(rslist_fujia("pagevol"))
		end if
		
		response.write rslist_fujia("pagevol")
		response.write "P"
		response.write "</td></tr>"
		rslist_yunyong.close()
		rslist_fujia.movenext
	loop
	rslist_fujia.close
	set rslist_fujia = nothing
		%>
    </table></td>
    <td><table width="95%"  border="0" cellspacing="0" cellpadding="0">
      <%
	if not isnull(rs("yunyong")) and rs("yunyong")<>"" then
		arr_idlist=split(rs("yunyong"),", ")
		arr_sllist=split(rs("sl"),", ")
	  count11=ubound(arr_idlist)+1
	  count22=0
	  for yy=1 to count11
		
		set rslistflag = conn.execute("select [type4] from yunyong where id="&arr_idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("type4")=1 then
				count22=count22+1
	%>
      <tr>
        <td><%
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&arr_idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_fd_name(arr_idlist(yy-1))<>"" then
			dict_fd_vol(arr_idlist(yy-1))=dict_fd_vol(arr_idlist(yy-1))+cint(arr_sllist(yy-1))
		else
			dict_fd_name(arr_idlist(yy-1))=rslist_yunyong("yunyong")
			dict_fd_vol(arr_idlist(yy-1))=cint(arr_sllist(yy-1))
		end if

		response.write arr_sllist(yy-1)
		response.write "��"
		rslist_yunyong.close()
		%></td>
      </tr>
      <%
			end if
			end if
			rslistflag.close()
		next
	end if
	
	set rslist_fujia = conn.execute("select fujia.jixiang,fujia.sl from fujia inner join yunyong on fujia.jixiang=yunyong.id where fujia.xiangmu_id="&rs("id")&" and yunyong.type4=1")
	do while not rslist_fujia.eof
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&rslist_fujia("jixiang"))
		response.Write "<tr><td>"&rslist_yunyong("yunyong")&"</td><td align=right>"
		
		if dict_fd_name(cstr(rslist_fujia("jixiang")))<>"" then
			dict_fd_vol(cstr(rslist_fujia("jixiang")))=dict_fd_vol(cstr(rslist_fujia("jixiang")))+cint(rslist_fujia("sl"))
		else
			dict_fd_name(cstr(rslist_fujia("jixiang")))=rslist_yunyong("yunyong")
			dict_fd_vol(cstr(rslist_fujia("jixiang")))=cint(rslist_fujia("sl"))
		end if
		
		response.write rslist_fujia("sl")
		response.write "��"
		response.write "</td></tr>"
		rslist_yunyong.close()
		rslist_fujia.movenext
	loop
	rslist_fujia.close
	set rslist_fujia = nothing
	%>
    </table></td>
    <td width="14%"><%if not isnull(rs("lc_wc")) then
		response.write datevalue(rs("lc_wc"))
	else
		response.write "&nbsp;"
	end if%></td>
  </tr>
  <%
 ' choucheng11=choucheng11+choucheng
   banmianll=banmianll+banmian
  fangdall=fangdall+fangda
 
  jixiang_money=jixiang_money+rs("jixiang_money")
  rs.movenext
  i=i+1
loop
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;����ϵ <%=formatnumber(jixiang_money,1,0,0,0)%>Ԫ&nbsp;&nbsp; �ܺ��� <%=formatnumber(allhqmoney,1,0,0,0)%>Ԫ&nbsp;&nbsp; ������ϵ��Ƭ������<%=allxpnum%> ��&nbsp;&nbsp;&nbsp; ������˴�����<%=xpcount%> ��<br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#FFFFFF">
          <td width="100" valign="top">&nbsp;��Ƭ��Ŀ�б�<br>
            &nbsp;���ܼ�<span id="sp_gp3">0</span>����</td>
          <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <%
	  if idlist="" or isnull(idlist) then
	  	response.write "<td>��</td>"
	  else
	  set rs_dg=server.createobject("adodb.recordset")
	  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&idlist&") and jixiang in (select id from yunyong where isgp=1) group by jixiang"
	  rs_dg.open sql,conn,1,1
	  gpvol=0
	  if not rs_dg.eof then
	  For i=1 to rs_dg.recordcount 
	  If rs_dg.eof Then Exit For
	  gpvol=gpvol+rs_dg("all_sl")
	  %>
              <td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%>��</td>
              <%
		if i mod 5=0 then
			response.write "</tr><tr>"
		end if
		rs_dg.Movenext
	next
	end if
	response.write "<script language='javascript'>document.getElementById('sp_gp3').innerHTML='"&gpvol&"';</script>"
	rs_dg.close
	set rs_dg=nothing
	end if
    %>
            </tr>
          </table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#FFFFFF">
          <td width="100" valign="top">&nbsp;�����Ŀ�б�<br>
            &nbsp;���ܼ�<span id="sp_xc3">0</span>P��</td>
          <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <%
		i=0
	  xcvol=0
	  if dict_xc_name.Count>0 then
	  	for each idno in dict_xc_name
	  %>
              <td><%=dict_xc_name(idno)%>:&nbsp;<%=dict_xc_vol(idno)%>P</td>
              <%
			i=i+1
			xcvol=xcvol+cint(dict_xc_vol(idno))
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
	  else
	    response.write "<td>��</td>"
      end if
	  response.write "<script language='javascript'>document.getElementById('sp_xc3').innerHTML='"&xcvol&"';</script>"
	set dict_xc_name=nothing
	set dict_xc_vol=nothing
    %>
            </tr>
          </table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#FFFFFF">
          <td width="100" valign="top">&nbsp;�Ŵ���Ŀ�б�<br>
            &nbsp;���ܼ�<span id="sp_fd3">0</span>�ţ�</td>
          <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <%
		i=0
	  fdvol=0
	  if dict_fd_name.Count>0 then
	  	for each idno in dict_fd_name
	  %>
              <td><%=dict_fd_name(idno)%>:&nbsp;<%=dict_fd_vol(idno)%>��</td>
              <%
			i=i+1
			fdvol=fdvol+cint(dict_fd_vol(idno))
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
	  else
	    response.write "<td>��</td>"
      end if
	  response.write "<script language='javascript'>document.getElementById('sp_fd3').innerHTML='"&fdvol&"';</script>"
	set dict_fd_name=nothing
	set dict_fd_vol=nothing
    %>
            </tr>
          </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<%end if%>
<%Call showYxTable()%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;<%'num11=conn.execute("select sum(tsvolume) from shejixiadan where "&GetSqlCheckDateString("lc_xp")&" and xp_name='"&conn.execute("select peplename from yuangong where username='"&userid&"'")(0)&"'")(0)
	'if isnull(num11) then num11=0
	'num12=conn.execute("select count(*) from shejixiadan where "&GetSqlCheckDateString("lc_xp2")&" and xp2_name='"&conn.execute("select peplename from yuangong where username='"&userid&"'")(0)&"'")(0)
	'if isnull(num12) then num12=0
	num13=conn.execute("select count(*) from shejixiadan where "&GetSqlCheckDateString("lc_sc")&" and sc_name='"&conn.execute("select peplename from yuangong where username='"&userid&"'")(0)&"'")(0)
	if isnull(num13) then num13=0
	num14=conn.execute("select count(*) from shejixiadan where "&GetSqlCheckDateString("lc_zd")&" and zd_name='"&conn.execute("select peplename from yuangong where username='"&userid&"'")(0)&"'")(0)
	if isnull(num14) then num14=0
	'response.Write num11
	%>�������: <%=num13%>��&nbsp;&nbsp;&nbsp;&nbsp;��Ʒ�������:<%=num14%>��<br>
&nbsp;���¹���:
<%
if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
	else
		gongzi=0
	end if
end if
if (fromtime<>"" and not isnull(fromtime)) and (totime<>"" and not isnull(totime)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
	else
		gongzi=0
	end if
end if

response.Write gongzi%>
Ԫ&nbsp;&nbsp;��ע:
<%if beizhu="" or isnull(beizhu) then 
response.Write "��"
else
response.Write beizhu
end if%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
Response.Write("&nbsp;ͶƱ��&nbsp;&nbsp;")
user_id = conn.execute("select id from yuangong where username='"&userid&"'")(0)

score=60
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��60��;&nbsp;&nbsp;"

score=80
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��80��;&nbsp;&nbsp;"

score=100
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��100��;&nbsp;&nbsp;"
%>
<br></td>
  </tr>
</table>
<%case 4
set rs=server.CreateObject("adodb.recordset")
'��Ӱ��ѡƬ
'��Ӱ�����½ɺ��ڿ�
rs.open "select * from shejixiadan where (cp_name='"&cur_peplename&"' or cp_name2='"&cur_peplename&"' or cp_name3='"&cur_peplename&"' or cp_name4='"&cur_peplename&"' or cp_name5='"&cur_peplename&"') and "&GetSqlCheckDateString("lc_ky"),conn,1,1

'(id in (select xiangmu_id from save_money where "&GetSqlCheckDateString("times")&") or id in (select xiangmu_id from sell_jilu where "&GetSqlCheckDateString("times")&")) and 

alldgmoney=0
cur_dgmoney = 0
allhqmoney=0
allhqqianmoney=0
hsky_vol=0
qtky_vol=0
allsavemoney=0
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">��ϵ/Ԫ</td>
    <td align="center">�ܺ���/Ƿ��</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">����/ǰ��/����</td>
    <td align="center">���</td>
    <td align="center">��ϵ����</td>
    <td align="center">������Ƭ</td>
    <td align="center">��Ӱ����</td>
    <td align="center">ǩ�����</td>
  </tr>
  <%do while not rs.eof
  		set rskyx = conn.execute("select * from jixiang where id="&rs("jixiang"))
  		if not (rskyx.eof and rskyx.bof) then
			if rskyx("type")=25 then
				hsky_vol = hsky_vol + 1
			else
				qtky_vol = qtky_vol + 1
			end if
		end if
		rskyx.close
		set rskyx = nothing
		
		num111=0
		if (not isnull(rs("cp_name")) and rs("cp_name")<>"") then num111=num111+1
		if (not isnull(rs("cp_name2")) and rs("cp_name2")<>"") then num111=num111+1
		if (not isnull(rs("cp_name3")) and rs("cp_name3")<>"") then num111=num111+1
		if (not isnull(rs("cp_name4")) and rs("cp_name4")<>"") then num111=num111+1
		if (not isnull(rs("cp_name5")) and rs("cp_name5")<>"") then num111=num111+1
		
		taoxi_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and [type]=1")(0)
		if isnull(taoxi_save) then taoxi_save=0
		fujia_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and [type]=2")(0)
		if isnull(fujia_save) then fujia_save=0
		fujia2_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&"  and "&GetSqlCheckDateString("times")&"and [type]=3")(0)
		if isnull(fujia2_save) then fujia2_save=0
		goumai_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and [type]=4")(0)
		if isnull(goumai_save) then goumai_save=0
		allsavemoney = allsavemoney + taoxi_save + fujia_save + fujia2_save + goumai_save
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
	msidlist=msidlist &", "& rs("id")
	%>    </td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"> 
      <%
	jixiang_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=1 and xiangmu_id="&rs("id")&"")(0)
	if isnull(jixiang_save) then jixiang_save=0
	response.Write rs("jixiang_money")%>    </td>
    <td align="center">
    <%
  	hq_fujia=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id"))(0)
	  if isnull(hq_fujia) then hq_fujia=0
  
  	'�����ܺ��ڽɿ�
  	hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
	if isnull(hq_savemoney) then hq_savemoney=0
	
	  'allhqmoney=allhqmoney+hq_fujia
	  
	  fujia_hepai = fujia_hepai + hq_fujia
	  if num111=1 then
	  	fujia_fenpai1 = fujia_fenpai1 + hq_fujia
	  else
	  	fujia_fenpai2 = fujia_fenpai2 + hq_fujia/num111
	  end if
	  
	  allhqqianmoney = allhqqianmoney + hq_fujia - hq_savemoney
	  response.Write round(hq_fujia/num111,1)&"/"&round((GetNonSaveMoney(rs("id"),2))/num111,1)%></td>
    <td align="center"><%
	all_wedvol = 0
	
	if rs("cp_name")<>"" and not isnull(rs("cp_name")) then
		response.write rs("cp_name")&"/"&rs("cp_wedvol")
		all_wedvol=all_wedvol+rs("cp_wedvol")
		if cur_peplename=rs("cp_name") then my_wedvol=rs("cp_wedvol")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name2")<>"" and not isnull(rs("cp_name2")) then
		response.write rs("cp_name2")&"/"&rs("cp_wedvol2")
		all_wedvol=all_wedvol+rs("cp_wedvol2")
		if cur_peplename=rs("cp_name2") then my_wedvol=rs("cp_wedvol2")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name3")<>"" and not isnull(rs("cp_name3")) then
		response.write rs("cp_name3")&"/"&rs("cp_wedvol3")
		all_wedvol=all_wedvol+rs("cp_wedvol3")
		if cur_peplename=rs("cp_name3") then my_wedvol=rs("cp_wedvol3")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name4")<>"" and not isnull(rs("cp_name4")) then
		response.write rs("cp_name4")&"/"&rs("cp_wedvol4")
		all_wedvol=all_wedvol+rs("cp_wedvol4")
		if cur_peplename=rs("cp_name4") then my_wedvol=rs("cp_wedvol4")
	else
		response.write "&nbsp;"
	end if
	if rs("cp_name5")<>"" and not isnull(rs("cp_name5")) then
		all_wedvol=all_wedvol+rs("cp_wedvol5")
		if cur_peplename=rs("cp_name5") then my_wedvol=rs("cp_wedvol5")
	end if
	'all_tx_wed=all_tx_wed+my_wedvol
	%></td>
    <td align="center"><%
	dgmoney=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	if isnull(dgmoney) then dgmoney=0
	alldgmoney=alldgmoney+dgmoney
	if my_wedvol="" or isnull(my_wedvol) then my_wedvol=0
	if hq_fujia="" or isnull(hq_fujia) then hq_fujia=0
	if all_wedvol=0 then
		response.write "0%/0/0"
	else
		per = round(my_wedvol/all_wedvol,2)
		hqs = per*100&"%/"&per*cint(hq_fujia)&"/"&per*cint(rs("jixiang_money"))
		response.write hqs
		allpersonhq = allpersonhq + per*cint(hq_fujia)
		cur_dgmoney = cur_dgmoney + per*dgmoney
	end if
	%></td>
    <td align="center"><%=GetWedVol(rs("id"))%></td>
    <td align="center"><%response.write rs("sl2")
	all_txVolume = all_txVolume + rs("sl2")
	%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
        <tr>
          <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
          <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
        </tr>
        <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
    </table></td>
    <td align="center"><%response.write rs("cpVolume")
	all_cpVolume = all_cpVolume + rs("cpVolume")
	%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rslf = server.CreateObject("adodb.recordset")
	rslf.open "SELECT hs_signtype.title, hs_signhistory.vol FROM hs_signtype INNER JOIN hs_signhistory ON hs_signtype.ID = hs_signhistory.typeid where hs_signhistory.userid="&cur_userid&" and hs_signhistory.xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("hs_signhistory.idate"),conn,1,1
	do while not rslf.eof
	%>
      <tr>
        <td>&nbsp;<%=rslf("title")%></td>
        <td align="right"><%=rslf("vol")%>&nbsp;</td>
      </tr>
      <%
		rslf.movenext
	loop
	rslf.close
	set rslf=nothing
	%>
    </table></td>
  </tr>
  <%
	fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=2 and xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia_save) then fujia_save=0
	
	'������º����տ�
	'response.write "����/"&rs("id")&"&nbsp;&nbsp;�ͻ�/"&conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)&"&nbsp;&nbsp;�����տ�/"&fujia_save&"<br>"
	  
	'num111=conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=4")(0)
	money13=conn.execute("select sum(dj*sl) from sell_jilu where "&GetSqlCheckDateString("times"))(0)
	money13=money13/num111
	if isnull(money13) then money13=0
	money13=formatnumber(money13,1,0,0,0)
	'fujia_save11=cint(fujia_save11+fujia_save/num111)

	fujia_save11=fujia_save11+fujia_save
	if num111=1 then
	  	hqsave_hepai1 = hqsave_hepai1 + fujia_save
	else
	  	hqsave_hepai2 = hqsave_hepai2 + fujia_save/num111
	end if
	
	'jixiang_money=clng(jixiang_money+rs("jixiang_money")/num111)
	jixiang_money=clng(jixiang_money+rs("jixiang_money"))
	money113=clng(money113+money13)
	sl2 = sl2 + rs("sl2")
	if idlist="" or isnull(idlist) then
		idlist = rs("id")
	else
		idlist = idlist & ", " & rs("id")
	end if
	rs.movenext
	i=i+1
loop
if msidlist<>"" then msidlist=mid(msidlist,3)
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;
      <%'response.Write formatnumber(allsavemoney,1,0,0,0)
'	  set rssyhq = conn.execute("select * from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and (j.cp_name='"&cur_peplename&"' or j.cp_name2='"&cur_peplename&"' or j.cp_name3='"&cur_peplename&"' or j.cp_name4='"&cur_peplename&"' or j.cp_name5='"&cur_peplename&"')  and "&GetSqlCheckDateString("s.times")&" and s.xiangmu_id not in (select id from shejixiadan where "&GetSqlCheckDateString("lc_cp")&")")
'	  response.write "���ڲ��<br>"
'	  do while not rssyhq.eof
'	  	response.write "����/"&rssyhq("xiangmu_id")&"&nbsp;&nbsp;�ͻ�/"&conn.execute("select lxpeple from kehu where id="&conn.execute("select kehu_id from shejixiadan where id="&rssyhq("xiangmu_id"))(0))(0)&"&nbsp;&nbsp;�����տ�/"&rssyhq("money")&"<br>"
'	  	rssyhq.movenext
'	  loop
'	  rssyhq.close
'	  set rssyhq = nothing
	  
	  'hqbk_money = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and (j.cp_name='"&cur_peplename&"' or j.cp_name2='"&cur_peplename&"' or j.cp_name3='"&cur_peplename&"' or j.cp_name4='"&cur_peplename&"' or j.cp_name5='"&cur_peplename&"')  and "&GetSqlCheckDateString("s.times")&" and s.xiangmu_id not in (select id from shejixiadan where "&GetSqlCheckDateString("lc_cp")&")")(0)
	  'hqbk_money = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and (j.cp_name='"&cur_peplename&"' or j.cp_name2='"&cur_peplename&"' or j.cp_name3='"&cur_peplename&"' or j.cp_name4='"&cur_peplename&"' or j.cp_name5='"&cur_peplename&"')  and "&GetSqlCheckDateString("s.times")&" and not ("&GetSqlCheckDateString("lc_cp")&")")(0)
	  'if isnull(hqbk_money) then hqbk_money  = 0
	  'response.Write formatnumber(fujia_save11,1,0,0,0)'&" + "& hqbk_money &" (���ڲ���)"%>
��ϵ��<%response.Write int(jixiang_money)
	jixiang_choucheng=int(jixiang_money)*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
	%>
    Ԫ&nbsp; &nbsp;�ܺ���(������)��<%response.Write fujia_hepai%> Ԫ&nbsp;&nbsp;�ϰ�԰�ֿ���<%response.Write formatnumber(fujia_fenpai1,1,0,0,0) & " + " & formatnumber(fujia_fenpai2,1,0,0,0)%> Ԫ&nbsp;&nbsp;���º����տ<%=formatnumber(hqsave_hepai1+hqsave_hepai2,1,0,0,0)%> Ԫ
    <%
	  flag2 = conn.execute("select scInvis from sysconfig")(0)
	  if flag2=1 then
	  %>
    ���˺��ڣ�
    <%response.Write allpersonhq%>
Ԫ &nbsp;<span class="STYLE10" style="display:none">��1��1�����Զ�&nbsp; ��Ƭ�ܽ��:
<%
	response.write alldgmoney
	'response.write money113
	'daogou_choucheng=money113*conn.execute("select choucheng5 from yuangong where username='"&userid&"'")(0)
  if isnull(jixiang_choucheng) then jixiang_choucheng=0
  if isnull(fujia_choucheng) then fujia_choucheng=0
  if isnull(daogou_choucheng) then  daogou_choucheng=0
	%>
Ԫ&nbsp; ����: <%=cur_dgmoney%>Ԫ�� </span>&nbsp;<%end if%>
����δ���ѣ�
      <%set rs_ds1 = server.createobject("adodb.recordset")
		set rs_ds3 = server.createobject("adodb.recordset")
		ds1_all = 0
		ds3_all = 0
		rs_ds1.open "select distinct s.id from shejixiadan s inner join kehu k on s.kehu_id=k.id where (cp_name='"&cur_peplename&"' or cp_name2='"&cur_peplename&"' or cp_name3='"&cur_peplename&"' or cp_name4='"&cur_peplename&"' or cp_name5='"&cur_peplename&"') and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		if not (rs_ds1.eof and rs_ds1.bof) then
			ds1_all = rs_ds1.recordcount
		else
			ds1_all = 0
		end if
		rs_ds1.close
		
		rs_ds3.open "select distinct s.id from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where (cp_name='"&cur_peplename&"' or cp_name2='"&cur_peplename&"' or cp_name3='"&cur_peplename&"' or cp_name4='"&cur_peplename&"' or cp_name5='"&cur_peplename&"') and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		if not (rs_ds3.eof and rs_ds3.bof) then
			ds3_all = rs_ds3.recordcount
		else
			ds3_all = 0
		end if
		rs_ds3.close
		
		response.write ds1_all-ds3_all&"��"%><br>&nbsp;����Ӱ:
      <%num12=conn.execute("select count(*) from shejixiadan where "&GetSqlCheckDateString("lc_xp")&" and xp_name='"&cur_peplename&"'")(0)
	if isnull(num12) then num12=0
	
	num11=conn.execute("select count(*) from shejixiadan where (cp_name='"&cur_peplename&"' or cp_name2='"&cur_peplename&"' or cp_name3='"&cur_peplename&"' or cp_name4='"&cur_peplename&"' or cp_name5='"&cur_peplename&"') and "&GetSqlCheckDateString("lc_cp"))(0)
	if isnull(num11) then num11=0
	response.Write num11
	%>
      ��&nbsp; ����ɫ:<%=num12%>
	��&nbsp;&nbsp;&nbsp;��ɴѡƬ
    <%=hsky_vol%>
    &nbsp;&nbsp;&nbsp; ����ѡƬ
    <%=qtky_vol%>    &nbsp; ����:
    <%if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
	else
		gongzi=0
	end if
end if
if (fromtime<>"" and not isnull(fromtime)) and (totime<>"" and not isnull(totime)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
	else
		gongzi=0
	end if
end if
response.Write gongzi%>
Ԫ&nbsp;&nbsp;��ע:
<%
if beizhu="" or isnull(beizhu) then 
response.Write "��"
else
response.Write beizhu
end if
%><br><%
		set rs_ds1 = server.createobject("adodb.recordset")
		set rs_ds2 = server.createobject("adodb.recordset")
		set rs_ds3 = server.createobject("adodb.recordset")
		
		rs_ds1.open "select distinct s.id from shejixiadan s inner join kehu k on s.kehu_id=k.id where s.cp_name='"&peplename&"' and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		if not (rs_ds1.eof and rs_ds1.bof) then
			ds1_all = rs_ds1.recordcount
		else
			ds1_all = 0
		end if
		rs_ds1.close
		
		rs_ds3.open "select distinct s.id from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where s.cp_name='"&peplename&"' and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		if not (rs_ds3.eof and rs_ds3.bof) then
			ds3_all = rs_ds3.recordcount
		else
			ds3_all = 0
		end if
		rs_ds3.close
		
		ds2_all = 0
		rs_ds2.open "select s.cp_name,f.money from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where s.cp_name='"&peplename&"' and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		do while not rs_ds2.eof
			ds2_all = ds2_all + rs_ds2("money")
			rs_ds2.movenext
		loop
		rs_ds2.close
		
		ds_count=0		'����
		ds1_count=0		'ѡƬ��¼����
		ds2_count=0		'ѡƬ���Ѻϼ�
		ds3_count=0		'ѡ�����Ѽ�¼����
		set rslost = conn.execute("select * from CustomerLostType order by px")
		do while not rslost.eof
			ds1 = 0
			ds2 = 0
			ds3 = 0
			
			rs_ds1.open "select distinct s.id from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.CustomerLostType="&rslost("id")&" and s.cp_name='"&peplename&"' and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			if not (rs_ds1.eof and rs_ds1.bof) then
				ds1 = rs_ds1.recordcount
			else
				ds1 = 0
			end if
			rs_ds1.close
			
			rs_ds3.open "select distinct s.id from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where k.CustomerLostType="&rslost("id")&" and s.cp_name='"&peplename&"' and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			if not (rs_ds3.eof and rs_ds3.bof) then
				ds3 = rs_ds3.recordcount
			else
				ds3 = 0
			end if
			rs_ds3.close
			
			rs_ds2.open "select s.ky_name,s.ky_name2,f.money from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where k.CustomerLostType="&rslost("id")&" and s.cp_name='"&peplename&"' and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			do while not rs_ds2.eof
				ds2 = ds2 + rs_ds2("money")
				rs_ds2.movenext
			loop
			rs_ds2.close
			
			ds1_count = ds1_count + ds1
			ds2_count = ds2_count + ds2
			ds3_count = ds3_count + ds3
			response.write rslost("title")&"ѡƬ"&ds1&"�� "
			response.write "δ����"& ds1-ds3 &"�� "
			response.write "��"&ds2&"Ԫ ƽ�����"
			if ds1=0 then 
				response.write ".0"
			else
				response.write formatnumber(ds2/ds1,1,0,0,0)
			end if
			response.write " Ԫ&nbsp;&nbsp;&nbsp;"
			ds_count = ds_count + 1
			if ds_count mod 2 = 0 then response.write "<br>&nbsp;"
			rslost.movenext
		loop
		rslost.close
		set rslost = nothing
		
		response.write "����ѡƬ"&ds1_all-ds1_count&"�� "
		response.write "δ����"& (ds1_all-ds3_all)-(ds1_count-ds3_count) &"�� "
		response.write "��"& ds2_all-ds2_count &"Ԫ ƽ�����"
		if (ds1_all-ds3_all)-(ds1_count-ds3_count)=0 then 
			response.write ".0"
		else
			response.write formatnumber((ds2_all-ds2_count)/(ds1_all-ds1_count),1,0,0,0)
		end if
		response.write " Ԫ"
%>
<br>
&nbsp;��ϵ������ <%=all_txVolume%> ��   &nbsp; ��Ӱ������ <%=all_cpVolume%> �ţ���ɫʦǩ��������<%signwedlist = ShowWedSignStats(msidlist, cur_userid)
if signwedlist<>"" then response.write "<br>&nbsp;ǩ�������"&signwedlist%><br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="15%" valign="top">&nbsp;���ڹ�Ƭ��</td>
	<td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
      <tr>
	 <%
	 if idlist<>"" then
	  set rs_dg=server.createobject("adodb.recordset")
	  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&idlist&") and jixiang in (select id from yunyong where isgp=1) and "&GetSqlCheckDateString("times")&" group by jixiang"
	  rs_dg.open sql,conn,1,1
	  if not rs_dg.eof then
	  For i=1 to rs_dg.recordcount 
	  If rs_dg.eof Then Exit For
	  %>
    <td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%> ��</td>
    <%
	if i mod 5=0 then
	response.write "</tr><tr>"
	end if
	rs_dg.Movenext
	next
	end if
	rs_dg.close
	set rs_dg=nothing
	end if
    %>
      </tr>
    </table></td>
    </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="15%" valign="top">&nbsp;��������������</td>
    <td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
	  set tonglei_rs=server.createobject("adodb.recordset")
	  'sql="Select name,sum(sl) as shuliang From sell_jilu group by name"
	  sql="Select name,sum(sl) as shuliang From sell_jilu where yuangong_id="&conn.execute("select id from yuangong where username='"&userid&"'")(0)&" and "&GetSqlCheckDateString("times")&" group by name"
	  tonglei_rs.open sql,conn,1,1
	  if not tonglei_rs.eof then
	  For i=1 to tonglei_rs.recordcount 
	  If tonglei_rs.eof Then Exit For
	  %>
        <td><%=tonglei_rs("name")%>:&nbsp;<%=tonglei_rs("shuliang")%> ��</td>
        <%
	if i mod 5=0 then
	response.write "</tr><tr>"
	end if
	tonglei_rs.Movenext
	next
	end if
	tonglei_rs.close
	set tonglei_rs=nothing
	rs.close
    %>
      </tr>
    </table></td>
  </tr>
</table></td>
  </tr>
</table>
<%
call init_key()
msidlist=""
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where (cp_name='"&cur_peplename&"' or cp_name2='"&cur_peplename&"' or cp_name3='"&cur_peplename&"' or cp_name4='"&cur_peplename&"' or cp_name5='"&cur_peplename&"') and "&GetSqlCheckDateString("lc_cp"),conn,1,1
%>
<div align="center" style="line-height:30px"> 
  <%response.write datearea%>
&nbsp; �����б�</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">��ϵ/Ԫ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">����/ǰ��/����</td>
    <td align="center">���</td>
    <td align="center">��ϵ����</td>
    <td align="center">��Ӱ����</td>
    <td align="center">ǩ�����</td>
  </tr>
  <%do while not rs.eof
  		set rskyx = conn.execute("select * from jixiang where id="&rs("jixiang"))
  		if not (rskyx.eof and rskyx.bof) then
			if rskyx("type")=25 then
				hsky_vol = hsky_vol + 1
			else
				qtky_vol = qtky_vol + 1
			end if
		end if
		rskyx.close
		set rskyx = nothing
		
		num111=0
		if (not isnull(rs("cp_name")) and rs("cp_name")<>"") then num111=num111+1
		if (not isnull(rs("cp_name2")) and rs("cp_name2")<>"") then num111=num111+1
		if (not isnull(rs("cp_name3")) and rs("cp_name3")<>"") then num111=num111+1
		if (not isnull(rs("cp_name4")) and rs("cp_name4")<>"") then num111=num111+1
		if (not isnull(rs("cp_name5")) and rs("cp_name5")<>"") then num111=num111+1
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
	msidlist = msidlist & ", " & rs("id")
	%>    </td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"><%
	response.Write rs("jixiang_money")
	AllXiangmuMoney = AllXiangmuMoney + rs("jixiang_money")%></td>
    <td align="center"><%
	all_wedvol = 0
	
	if rs("cp_name")<>"" and not isnull(rs("cp_name")) then
		response.write rs("cp_name")&"/"&rs("cp_wedvol")
		all_wedvol=all_wedvol+rs("cp_wedvol")
		if cur_peplename=rs("cp_name") then my_wedvol=rs("cp_wedvol")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name2")<>"" and not isnull(rs("cp_name2")) then
		response.write rs("cp_name2")&"/"&rs("cp_wedvol2")
		all_wedvol=all_wedvol+rs("cp_wedvol2")
		if cur_peplename=rs("cp_name2") then my_wedvol=rs("cp_wedvol2")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name3")<>"" and not isnull(rs("cp_name3")) then
		response.write rs("cp_name3")&"/"&rs("cp_wedvol3")
		all_wedvol=all_wedvol+rs("cp_wedvol3")
		if cur_peplename=rs("cp_name3") then my_wedvol=rs("cp_wedvol3")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name4")<>"" and not isnull(rs("cp_name4")) then
		response.write rs("cp_name4")&"/"&rs("cp_wedvol4")
		all_wedvol=all_wedvol+rs("cp_wedvol4")
		if cur_peplename=rs("cp_name4") then my_wedvol=rs("cp_wedvol4")
	else
		response.write "&nbsp;"
	end if
	if rs("cp_name5")<>"" and not isnull(rs("cp_name5")) then
		all_wedvol=all_wedvol+rs("cp_wedvol5")
		if cur_peplename=rs("cp_name5") then my_wedvol=rs("cp_wedvol5")
	end if
	'all_tx_wed=all_tx_wed+my_wedvol
	%></td>
    <td align="center"><%
	dgmoney=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	if isnull(dgmoney) then dgmoney=0
	alldgmoney=alldgmoney+dgmoney
	if my_wedvol="" or isnull(my_wedvol) then my_wedvol=0
	if hq_fujia="" or isnull(hq_fujia) then hq_fujia=0
	if all_wedvol=0 then
		response.write "0%/0/0"
	else
		per = round(my_wedvol/all_wedvol,2)
		hqs = per*100&"%/"&per*cint(hq_fujia)&"/"&per*cint(rs("jixiang_money"))
		response.write hqs
		allpersonhq = allpersonhq + per*cint(hq_fujia)
		cur_dgmoney = cur_dgmoney + per*dgmoney
	end if
	%></td>
    <td align="center"><%=GetWedVol(rs("id"))%></td>
    <td align="center"><%response.write rs("sl2")
	all_txVolume = all_txVolume + rs("sl2")
	%></td>
    <td align="center"><%response.write rs("cpVolume")
	all_cpVolume = all_cpVolume + rs("cpVolume")
	%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rslf = server.CreateObject("adodb.recordset")
	rslf.open "SELECT hs_signtype.title, hs_signhistory.vol FROM hs_signtype INNER JOIN hs_signhistory ON hs_signtype.ID = hs_signhistory.typeid where hs_signhistory.userid="&cur_userid&" and hs_signhistory.xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("hs_signhistory.idate"),conn,1,1
	do while not rslf.eof
	%>
      <tr>
        <td>&nbsp;<%=rslf("title")%></td>
        <td align="right"><%=rslf("vol")%>&nbsp;</td>
      </tr>
      <%
		rslf.movenext
	loop
	rslf.close
	set rslf=nothing
	%>
    </table></td>
  </tr>
  <%
	fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=2 and xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia_save) then fujia_save=0
	
	'������º����տ�
	'response.write "����/"&rs("id")&"&nbsp;&nbsp;�ͻ�/"&conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)&"&nbsp;&nbsp;�����տ�/"&fujia_save&"<br>"
	  
	'num111=conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=4")(0)
	money13=conn.execute("select sum(dj*sl) from sell_jilu where "&GetSqlCheckDateString("times"))(0)
	money13=money13/num111
	if isnull(money13) then money13=0
	money13=formatnumber(money13,1,0,0,0)
	'fujia_save11=cint(fujia_save11+fujia_save/num111)

	fujia_save11=fujia_save11+fujia_save
	if num111=1 then
	  	hqsave_hepai1 = hqsave_hepai1 + fujia_save
	else
	  	hqsave_hepai2 = hqsave_hepai2 + fujia_save/num111
	end if
	
	'jixiang_money=clng(jixiang_money+rs("jixiang_money")/num111)
	jixiang_money=clng(jixiang_money+rs("jixiang_money"))
	money113=clng(money113+money13)
	sl2 = sl2 + rs("sl2")
	if idlist="" or isnull(idlist) then
		idlist = rs("id")
	else
		idlist = idlist & ", " & rs("id")
	end if
	rs.movenext
	i=i+1
loop
rs.close
if msidlist<>"" then msidlist=mid(msidlist,3)
  %>
</table>
<table width="100%"  border="0" cellpadding="0" cellspacing="0"><tr>
  <td>
 &nbsp;�ϼ���ϵ��� <%=formatnumber(AllXiangmuMoney,1,0,0,0)%> Ԫ<br>
&nbsp;��������Ӱ
<%
sycount=0
syall=conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where (s.cp_name='"&cur_peplename&"' or s.cp_name2='"&cur_peplename&"' or s.cp_name3='"&cur_peplename&"' or s.cp_name4='"&cur_peplename&"' or s.cp_name5='"&cur_peplename&"') and "&GetSqlCheckDateString("s.lc_cp"))(0)
if isnull(syall) then syall=0
response.write syall
%>�� (<%set rssy = conn.execute("select * from CustomerLostType order by px")
do while not rssy.eof
	sy = conn.execute("select count(*) from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.CustomerLostType="&rssy("id")&" and (s.cp_name='"&cur_peplename&"' or s.cp_name2='"&cur_peplename&"' or s.cp_name3='"&cur_peplename&"' or s.cp_name4='"&cur_peplename&"' or s.cp_name5='"&cur_peplename&"') and "&GetSqlCheckDateString("s.lc_cp"))(0)
	if isnull(sy) then sy=0
	sycount = sycount + sy
	response.write rssy("title")&sy&",&nbsp;"
	rssy.movenext
loop
rssy.close
set rssy = nothing
response.write "����" & syall - sycount
%>)<%signwedlist = ShowWedSignStats(msidlist, cur_userid)
if signwedlist<>"" then response.write "<br>&nbsp;ǩ�������"&signwedlist%></td>
</tr></table><br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" align="center"><%response.write datearea%>
      &nbsp;������ѡƬ������ϸ��</td>
  </tr>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" style="richness:1px">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">����ϵ��</td>
    <td align="center">��ϵ�ɷ�/(�Ŷ�)</td>
    <td width="16%" align="center">ѡƬ�����ܽ��</td>
    <td align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ƭ���͡�</td>
    <td align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
  </tr>
  <%
  Call init_key()
  rs.open "select * from shejixiadan where (cp_name='"&peplename&"' or cp_name2='"&peplename&"' or cp_name3='"&peplename&"' or cp_name4='"&peplename&"' or cp_name5='"&peplename&"') and "&sql_time&" and not ("&GetSqlCheckDateString("lc_cp")&") and id in (select xiangmu_id from save_money where [type]=2 and "&GetSqlCheckDateString("times")&")",conn,1,1

  'msidlist=","
  do while not rs.eof
  str_sm=""
  count111=0
  if not isnull(rs("cp_name")) and rs("cp_name")<>"" then count111=count111+1
  if not isnull(rs("cp_name2")) and rs("cp_name2")<>"" then count111=count111+1
  if not isnull(rs("cp_name3")) and rs("cp_name3")<>"" then count111=count111+1
  if not isnull(rs("cp_name4")) and rs("cp_name4")<>"" then count111=count111+1
  if not isnull(rs("cp_name5")) and rs("cp_name5")<>"" then count111=count111+1
  
  jixiang_money=jixiang_money+rs("jixiang_money")
  
  '�������½ɺ��ڿ�
  hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
  if count111=1 then
	  hqsave_hepai1 = hqsave_hepai1 + hq_indate_savemoney
  else
	  hqsave_hepai2 = hqsave_hepai2 + hq_indate_savemoney/count111
  end if

  	if isnull(money2) then money2=0
	sm2_money=money2
	hq_indate_savemoney=hq_indate_savemoney/count111
  
  '�����ܺ���
  hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
  if isnull(hq_money) then hq_money = 0
  
  '�����ܺ��ڽɿ�
  hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
  
  
  	if hq_money=hq_savemoney then
  		ReceivablesMoney = ReceivablesMoney + hq_money/count111
  	end if

  'if hq_money=hq_indate_savemoney then 
  '	RecFujiaMoney = RecFujiaMoney+hq_mymoney
	'AllRecFujiaMoney = AllRecFujiaMoney+hq_money
  'end if
  
  'hq_minesavemoney = conn.execute("select sum(money) from save_money where [type]=2 and userid='"&userid&"' and xiangmu_id="&rs("id"))(0)
  set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
  do while not rshq.eof
  	if rshq("userid")=userid or rshq("userid2")=userid then
	  if rshq("userid")<>"" and not isnull(rshq("userid2")) then
		hq_mymoney = hq_mymoney + rshq("money")/2
	  else
	  	hq_mymoney = hq_mymoney + rshq("money")
  	  end if
	end if
	rshq.movenext
  loop
  rshq.close
  set rshq=nothing
  
  if isnull(hq_savemoney) then hq_savemoney = 0
  
  '��Ƿ��
  hq_notsavemoney=hq_notsavemoney+hq_money-hq_savemoney
  
  '�ܺ���
  hq_allmoney=hq_allmoney+hq_money
  
  '�����ܺ��ڽɿ�
  hq_indate_allsavemoney=hq_indate_allsavemoney+hq_indate_savemoney
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		msidlist=msidlist&rs("id")&","
	%>
    </td>
    <td align="center"><%
	 response.Write conn.execute("select lxpeple from kehu where id="&rs("kehu_id"))(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>
    <td align="center"><% 
		jx_money = rs("jixiang_money")
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
    <td align="center"><%money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money1) then money1=0
	response.Write formatnumber(money1,1,0,0,0)
	%></td>
    <td align="center" bgcolor="#ffffff"><%
	hqallmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	if isnull(hqallmoney) then hqallmoney=0
	fujia_hepai = fujia_hepai + hqallmoney
	if count111=1 then
		fujia_fenpai1 = fujia_fenpai1 + hq_fujia
	else
	  	fujia_fenpai2 = fujia_fenpai2 + hq_fujia/count111
	end if
	response.write Formatnumber(hqallmoney/count111,1,0,0,0)
	%></td>
    <td align="center"><%
	money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id"))(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
	if rs("cp_name")<>cur_peplename then response.Write "/"&rs("cp_name")
	if rs("cp_name2")<>cur_peplename then response.Write "/"&rs("cp_name2")
	if rs("cp_name3")<>cur_peplename then response.Write "/"&rs("cp_name3")
	if rs("cp_name4")<>cur_peplename then response.Write "/"&rs("cp_name4")
	if rs("cp_name5")<>cur_peplename then response.Write "/"&rs("cp_name5")
	%></td>
    <td align="center" bgcolor="#ffffff"><%if rs("cp_name")<>cur_peplename and rs("cp_name2")<>cur_peplename and rs("cp_name3")<>cur_peplename and rs("cp_name4")<>cur_peplename and rs("cp_name5")<>cur_peplename then
		response.write "0"
	else%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
          <tr>
            <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
            <td>&nbsp;<%=rsdg("all_sl")%></td>
          </tr>
          <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
        </table>
      <%end if%></td>
    <td align="center"><%=GetNonSaveMoney(rs("id"),0)%></td>
  </tr>
  <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
 
  rs.movenext
  i=i+1
loop
rs.close()
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;��ϵ��<%response.Write int(jixiang_money)%> Ԫ&nbsp; &nbsp;�ܺ���(������)��<%response.Write fujia_hepai%> Ԫ&nbsp; &nbsp;�ϰ�԰�ֿ���<%response.Write formatnumber(fujia_fenpai1,1,0,0,0) & " + " & formatnumber(fujia_fenpai2,1,0,0,0)%> Ԫ&nbsp; &nbsp;ѡƬ���� <%=Formatnumber(hqsave_hepai1+ hqsave_hepai2,1,0,0,0)%> Ԫ&nbsp;&nbsp;&nbsp; &nbsp;�ۼƺ���ѡƬǷ��
      <%
	tmp_fujia_money = conn.execute("select sum(f.money) from fujia f inner join shejixiadan s on f.xiangmu_id=s.id where (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"')")(0)
	tmp_save_money = conn.execute("select sum(m.money) from save_money m inner join shejixiadan s on m.xiangmu_id=s.id where m.type=2 and (s.userid='"&userid&"' or s.userid2='"&userid&"' or s.userid3='"&userid&"')")(0)
	if isnull(tmp_fujia_money) then tmp_fujia_money = 0
	if isnull(tmp_save_money) then tmp_save_money = 0
	response.write Formatnumber(tmp_fujia_money-tmp_save_money,1,0,0,0)%>
      Ԫ&nbsp;&nbsp;&nbsp; &nbsp;������� <%=Formatnumber(ReceivablesMoney,1,0,0,0)%> Ԫ</td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><%Call showYxTable()%>
      <%
Call init_key()
set rs6=server.CreateObject("adodb.recordset")
rs6.open "select * from shejixiadan where xp_name='"&cur_peplename&"' and "&GetSqlCheckDateString("lc_ky"),conn,1,1
xpcount = rs6.recordcount
%>
      <div align="center" style="line-height:30px">
        <%response.write datearea%>
  &nbsp;
        ��ɫ����</div>
      <table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
        <tr bgcolor="#99FFFF">
          <td width="15%" height="19">&nbsp;&nbsp;����</td>
          <td width="18%" align="center">�ͻ�</td>
          <td align="center">������Ŀ</td>
          <td width="18%" align="center">��Ƭ���/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
          <td width="12%" align="center">��ɫ����</td>
          <td width="12%" align="center">��ϵ����</td>
        </tr>
        <%
  allxpnum=0
  alltsnum=0
  idlist=""
  do while not rs6.eof
  		allxpnum = allxpnum + rs6("sl2")
		alltsnum = alltsnum + rs6("tsvolume")
  %>
        <tr bgcolor="#FFFFFF">
          <td>&nbsp;
              <% response.Write rs6("id")
	if idlist="" or isnull(idlist) then
		idlist=rs6("id")
	else
		idlist=idlist&", "&rs6("id")
	end if
	%>          </td>
          <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs6("kehu_id")&"")(0)%></td>
          <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs6("id")&" and "&GetSqlCheckDateString("times")&" group by jixiang")
	do while not rsdg.eof
	%>
              <tr>
                <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
                <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
                <td>&nbsp;<%=rsdg("all_money")%>Ԫ&nbsp;</td>
              </tr>
              <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
          </table></td>
          <td align="center"><%
	  dgmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs6("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	  if isnull(dgmoney) then dgmoney=0
	  response.write dgmoney
	money13=conn.execute("select sum(dj*sl) from sell_jilu where "&GetSqlCheckDateString("times")&"")(0)
	if isnull(money13) then money13=0
	money13=formatnumber(money13,1,0,0,0)
	%>          </td>
          <td align="center"><%=rs6("tsVolume")%></td>
          <td align="center"><%=rs6("sl2")%></td>
        </tr>
        <%
    jixiang_money=jixiang_money+jixiang_save
	money113=money113+money13
	'fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=2 and xiangmu_id="&rs6("id")&"")(0)
	'if isnull(fujia_save) then fujia_save=0
	'fujia_save11=fujia_save11+fujia_save
  rs6.movenext
  i=i+1
loop

  %>
      </table>
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;���º����տ
            <%'response.Write formatnumber(allsavemoney,1,0,0,0)
	  fujia_save11 = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and j.xp_name='"&cur_peplename&"' and "&GetSqlCheckDateString("s.times")&" and "&GetSqlCheckDateString("j.lc_ky"))(0)
	  if isnull(fujia_save11) then fujia_save11=0
	  hqbk_money = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and xp_name='"&cur_peplename&"' and "&GetSqlCheckDateString("s.times")&" and s.xiangmu_id not in (select id from shejixiadan where "&GetSqlCheckDateString("lc_xp")&")")(0)
	  if isnull(hqbk_money) then hqbk_money  = 0
	  response.Write formatnumber(fujia_save11,1,0,0,0)&" + "& hqbk_money &" (���ڲ���)"%>
Ԫ&nbsp; &nbsp;������ϵ��Ƭ������<%=allxpnum%> ��&nbsp;&nbsp;&nbsp; ���µ�ɫ��ϵ������<%=alltsnum%> ��&nbsp;&nbsp;&nbsp; ������˴�����<%=xpcount%> ��</td>
        </tr>
      </table>
      <%call ShowSuitType(idlist)%>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#FFFFFF">
          <td width="15%" valign="top">&nbsp;��Ƭ��Ŀ�б�</td>
          <td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <%
	  if idlist="" or isnull(idlist) then
	  	response.write "<td>��</td>"
	  else
		  set rs_dg=server.createobject("adodb.recordset")
		  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&idlist&") and jixiang in (select id from yunyong where isgp=1) group by jixiang"
		  rs_dg.open sql,conn,1,1
		  if not rs_dg.eof then
		  For i=1 to rs_dg.recordcount 
		  If rs_dg.eof Then Exit For
		  %>
                <td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%> ��</td>
                <%
		if i mod 5=0 then
		response.write "</tr><tr>"
		end if
		rs_dg.Movenext
		next
		end if
		rs_dg.close
		set rs_dg=nothing
    end if%>
              </tr>
          </table></td>
        </tr>
      </table>
      <%
Call init_key()
if instr(qj_flag,"1")>0 then
	set rs=server.CreateObject("adodb.recordset")
	chk_peplename = conn.execute("select peplename from yuangong where username='"&userid&"'")(0)
	rs.open "select * from shejixiadan where (cp_name='"&chk_peplename&"' or cp_name2='"&chk_peplename&"' or cp_name3='"&chk_peplename&"' or cp_name4='"&chk_peplename&"' or cp_name5='"&chk_peplename&"') and "&GetSqlCheckDateString("lc_wc"),conn,1,1
%>
      <div align="center" style="line-height:30px"> 
        <%response.write datearea%>
&nbsp; ����ȡ���б�</div>
      <table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
        <tr bgcolor="#99FFFF">
          <td height="19" align="center">����</td>
          <td align="center">�ͻ�</td>
          <td align="center">��ϵ/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
          <td align="center">ѡƬ���/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
          <td align="center">������Ƭ</td>
          <td align="center">��Ӱ/��Ƭ</td>
          <td align="center">��Ӱ/��Ƭ</td>
          <td align="center">��Ӱ/��Ƭ</td>
          <td align="center">��Ӱ/��Ƭ</td>
          <td align="center">��Ӱ/��Ƭ</td>
          <td align="center">����/����</td>
          <td align="center">���</td>
        </tr>
        <%do while not rs.eof
  %>
        <tr bgcolor="#FFFFFF">
          <td align="center"><% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
	
	%></td>
          <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
          <td align="center"><%num=conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=4")(0)
	jixiang_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=1 and xiangmu_id="&rs("id")&"")(0)
	if isnull(jixiang_save) then jixiang_save=0
	response.Write rs("jixiang_money")
	jixiang_money=jixiang_money+rs("jixiang_money")%></td>
          <td align="center"><%
  	hq_fujia=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id"))(0)
	  if isnull(hq_fujia) then hq_fujia=0
	  allhqmoney=allhqmoney+hq_fujia
	  response.Write cint(hq_fujia)&"Ԫ"%></td>
          <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
            <tr>
              <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
              <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
            </tr>
            <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
          </table></td>
          <td align="center"><%
	all_wedvol = 0
	
	if rs("cp_name")<>"" and not isnull(rs("cp_name")) then
		response.write rs("cp_name")&"/"&rs("cp_wedvol")
		all_wedvol=all_wedvol+rs("cp_wedvol")
		if cur_peplename=rs("cp_name") then my_wedvol=rs("cp_wedvol")
	else
		response.write "&nbsp;"
	end if%></td>
          <td align="center"><%if rs("cp_name2")<>"" and not isnull(rs("cp_name2")) then
		response.write rs("cp_name2")&"/"&rs("cp_wedvol2")
		all_wedvol=all_wedvol+rs("cp_wedvol2")
		if cur_peplename=rs("cp_name2") then my_wedvol=rs("cp_wedvol2")
	else
		response.write "&nbsp;"
	end if%></td>
          <td align="center"><%if rs("cp_name3")<>"" and not isnull(rs("cp_name3")) then
		response.write rs("cp_name3")&"/"&rs("cp_wedvol3")
		all_wedvol=all_wedvol+rs("cp_wedvol3")
		if cur_peplename=rs("cp_name3") then my_wedvol=rs("cp_wedvol3")
	else
		response.write "&nbsp;"
	end if%></td>
          <td align="center"><%if rs("cp_name4")<>"" and not isnull(rs("cp_name4")) then
		response.write rs("cp_name4")&"/"&rs("cp_wedvol4")
		all_wedvol=all_wedvol+rs("cp_wedvol4")
		if cur_peplename=rs("cp_name4") then my_wedvol=rs("cp_wedvol4")
	else
		response.write "&nbsp;"
	end if%></td>
          <td align="center"><%if rs("cp_name5")<>"" and not isnull(rs("cp_name5")) then
		response.write rs("cp_name5")&"/"&rs("cp_wedvol5")
		all_wedvol=all_wedvol+rs("cp_wedvol5")
		if cur_peplename=rs("cp_name5") then my_wedvol=rs("cp_wedvol5")
	else
		response.write "&nbsp;"
	end if%></td>
          <td align="center"><%
	dgmoney=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	if isnull(dgmoney) then dgmoney=0
	alldgmoney=alldgmoney+dgmoney
	if all_wedvol=0 then
		response.write "0%/0"
	else
		per = round(my_wedvol/all_wedvol,2)
		response.write per*100&"%/"&per*cint(hq_fujia)
		allpersonhq = allpersonhq+per*100&"%/"&per*cint(hq_fujia)
		cur_dgmoney = cur_dgmoney + per*dgmoney
	end if
	all_tx_wed=all_tx_wed+my_wedvol
	%></td>
          <td align="center"><%=GetWedVol(rs("id"))%></td>
        </tr>
        <%
	fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=2 and xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia_save) then fujia_save=0
	  
	num111=conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=4")(0)
	money13=conn.execute("select sum(dj*sl) from sell_jilu where "&GetSqlCheckDateString("times")&"")(0)
	money13=money13/num111
	if isnull(money13) then money13=0
	money13=formatnumber(money13,1,0,0,0)
	fujia_save11=cint(fujia_save11+fujia_save/num)
	jixiang_money=clng(jixiang_money+jixiang_save/num)
	money113=money113+money13
	sl2 = sl2 + rs("sl2")
	if idlist="" or isnull(idlist) then
		idlist = rs("id")
	else
		idlist = idlist & ", " & rs("id")
	end if
	rs.movenext
	i=i+1
loop

  %>
      </table>
      &nbsp;��ϵ�ܽ�
      <%response.Write int(jixiang_money)
	jixiang_choucheng=int(jixiang_money)*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
	%>
Ԫ&nbsp; &nbsp;�ŶӺ��ڣ�
<%response.Write allhqmoney
	fujia_choucheng=allhqmoney*int(jixiang_money)*conn.execute("select choucheng2 from yuangong where username='"&userid&"'")(0)
	%>
Ԫ&nbsp;
<%
	  flag2 = conn.execute("select scInvis from sysconfig")(0)
	  if flag2=1 then
	  %>
���˺��ڣ�<%=allpersonhq%> Ԫ&nbsp;<span class="STYLE9">&nbsp;( 1��1���� ��Ƭ�ܽ��
<%
	response.write alldgmoney
	
	daogou_choucheng=money113*conn.execute("select choucheng5 from yuangong where username='"&userid&"'")(0)
  if isnull(jixiang_choucheng) then jixiang_choucheng=0
  if isnull(fujia_choucheng) then fujia_choucheng=0
  if isnull(daogou_choucheng) then  daogou_choucheng=0
	%>
Ԫ&nbsp; ���� <%=cur_dgmoney%>Ԫ) </span>&nbsp;&nbsp;&nbsp;
<%end if%>
<br>
&nbsp;�ϼ���ϵ������<%=sl2%> ��&nbsp;&nbsp;&nbsp;&nbsp; �ϼƺ���������<%=all_tx_wed%> ��
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="15%" valign="top">&nbsp;���ڹ�Ƭ�б�</td>
    <td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
				if idlist<>"" then
				  set rs_dg=server.createobject("adodb.recordset")
				  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&idlist&") and jixiang in (select id from yunyong where isgp=1) and "&GetSqlCheckDateString("times")&" group by jixiang"
				  rs_dg.open sql,conn,1,1
				  if not rs_dg.eof then
				  For i=1 to rs_dg.recordcount 
				  If rs_dg.eof Then Exit For
				  %>
        <td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%> ��</td>
        <%
				if i mod 5=0 then
				response.write "</tr><tr>"
				end if
				rs_dg.Movenext
				next
				end if
				rs_dg.close
				set rs_dg=nothing
				end if
				%>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="15%" valign="top">&nbsp;��������������</td>
    <td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
	  set tonglei_rs=server.createobject("adodb.recordset")
	  'sql="Select name,sum(sl) as shuliang From sell_jilu group by name"
	  sql="Select name,sum(sl) as shuliang From sell_jilu where yuangong_id="&conn.execute("select id from yuangong where username='"&userid&"'")(0)&" and "&GetSqlCheckDateString("times")&" group by name"
	  tonglei_rs.open sql,conn,1,1
	  if not tonglei_rs.eof then
	  For i=1 to tonglei_rs.recordcount 
	  If tonglei_rs.eof Then Exit For
	  %>
        <td><%=tonglei_rs("name")%>:&nbsp;<%=tonglei_rs("shuliang")%> ��</td>
        <%
	if i mod 5=0 then
	response.write "</tr><tr>"
	end if
	tonglei_rs.Movenext
	next
	end if
	tonglei_rs.close
	set tonglei_rs=nothing
    %>
      </tr>
    </table></td>
  </tr>
</table>
<%end if%>
<%
Call init_key()
if instr(qj_flag,"2")>0 then
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from shejixiadan where xp_name='"&peplename&"' and "&GetSqlCheckDateString("lc_wc"),conn,1,1
%>
<div align="center" style="line-height:30px"> 
  <%response.write datearea%>
&nbsp; ��ɫȡ���б�</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">��ϵ/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">ѡƬ���/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">������Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">��Ӱ/��Ƭ</td>
    <td align="center">����/����</td>
    <td align="center">���</td>
  </tr>
  <%do while not rs.eof
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
	
	%></td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"><%num=conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=4")(0)
	jixiang_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=1 and xiangmu_id="&rs("id")&"")(0)
	if isnull(jixiang_save) then jixiang_save=0
	response.Write rs("jixiang_money")
	jixiang_money=jixiang_money+rs("jixiang_money")%></td>
    <td align="center"><%
  	hq_fujia=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id"))(0)
	  if isnull(hq_fujia) then hq_fujia=0
	  allhqmoney=allhqmoney+hq_fujia
	  response.Write cint(hq_fujia)&"Ԫ"%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
      <tr>
        <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
        <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
      </tr>
      <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
    </table></td>
    <td align="center"><%
	all_wedvol = 0
	
	if rs("cp_name")<>"" and not isnull(rs("cp_name")) then
		response.write rs("cp_name")&"/"&rs("cp_wedvol")
		all_wedvol=all_wedvol+rs("cp_wedvol")
		if cur_peplename=rs("cp_name") then my_wedvol=rs("cp_wedvol")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name2")<>"" and not isnull(rs("cp_name2")) then
		response.write rs("cp_name2")&"/"&rs("cp_wedvol2")
		all_wedvol=all_wedvol+rs("cp_wedvol2")
		if cur_peplename=rs("cp_name2") then my_wedvol=rs("cp_wedvol2")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name3")<>"" and not isnull(rs("cp_name3")) then
		response.write rs("cp_name3")&"/"&rs("cp_wedvol3")
		all_wedvol=all_wedvol+rs("cp_wedvol3")
		if cur_peplename=rs("cp_name3") then my_wedvol=rs("cp_wedvol3")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name4")<>"" and not isnull(rs("cp_name4")) then
		response.write rs("cp_name4")&"/"&rs("cp_wedvol4")
		all_wedvol=all_wedvol+rs("cp_wedvol4")
		if cur_peplename=rs("cp_name4") then my_wedvol=rs("cp_wedvol4")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%if rs("cp_name5")<>"" and not isnull(rs("cp_name5")) then
		response.write rs("cp_name5")&"/"&rs("cp_wedvol5")
		all_wedvol=all_wedvol+rs("cp_wedvol5")
		if cur_peplename=rs("cp_name5") then my_wedvol=rs("cp_wedvol5")
	else
		response.write "&nbsp;"
	end if%></td>
    <td align="center"><%
	dgmoney=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	if isnull(dgmoney) then dgmoney=0
	alldgmoney=alldgmoney+dgmoney
	if all_wedvol=0 then
		response.write "0%/0"
	else
		per = round(my_wedvol/all_wedvol,2)
		response.write per*100&"%/"&per*cint(hq_fujia)
		allpersonhq = allpersonhq+per*100&"%/"&per*cint(hq_fujia)
		cur_dgmoney = cur_dgmoney + per*dgmoney
	end if
	all_tx_wed=all_tx_wed+my_wedvol
	%></td>
    <td align="center"><%=GetWedVol(rs("id"))%></td>
  </tr>
  <%
	fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=2 and xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia_save) then fujia_save=0
	  
	num111=conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=4")(0)
	money13=conn.execute("select sum(dj*sl) from sell_jilu where "&GetSqlCheckDateString("times")&"")(0)
	money13=money13/num111
	if isnull(money13) then money13=0
	money13=formatnumber(money13,1,0,0,0)
	fujia_save11=cint(fujia_save11+fujia_save/num)
	jixiang_money=clng(jixiang_money+jixiang_save/num)
	money113=money113+money13
	sl2 = sl2 + rs("sl2")
	if idlist="" or isnull(idlist) then
		idlist = rs("id")
	else
		idlist = idlist & ", " & rs("id")
	end if
	rs.movenext
	i=i+1
loop

  %>
</table>
&nbsp;��ϵ�ܽ�
<%response.Write int(jixiang_money)
	jixiang_choucheng=int(jixiang_money)*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
	%>
Ԫ&nbsp; &nbsp;�ŶӺ��ڣ�
<%response.Write allhqmoney
	fujia_choucheng=allhqmoney*int(jixiang_money)*conn.execute("select choucheng2 from yuangong where username='"&userid&"'")(0)
	%>
Ԫ&nbsp;
<%
	  flag2 = conn.execute("select scInvis from sysconfig")(0)
	  if flag2=1 then
	  %>
���˺��ڣ�<%=allpersonhq%> Ԫ&nbsp;<span class="STYLE9">&nbsp;( 1��1���� ��Ƭ�ܽ��
<%
	response.write alldgmoney
	
	daogou_choucheng=money113*conn.execute("select choucheng5 from yuangong where username='"&userid&"'")(0)
  if isnull(jixiang_choucheng) then jixiang_choucheng=0
  if isnull(fujia_choucheng) then fujia_choucheng=0
  if isnull(daogou_choucheng) then  daogou_choucheng=0
	%>
Ԫ&nbsp; ���� <%=cur_dgmoney%>Ԫ) </span>&nbsp;&nbsp;&nbsp;
<%end if%>
<br>
&nbsp;�ϼ���ϵ������<%=sl2%> ��&nbsp;&nbsp;&nbsp;&nbsp; �ϼƺ���������<%=all_tx_wed%> ��
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="15%" valign="top">&nbsp;���ڹ�Ƭ�б�</td>
    <td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
				if idlist<>"" then
				  set rs_dg=server.createobject("adodb.recordset")
				  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&idlist&") and jixiang in (select id from yunyong where isgp=1) and "&GetSqlCheckDateString("times")&" group by jixiang"
				  rs_dg.open sql,conn,1,1
				  if not rs_dg.eof then
				  For i=1 to rs_dg.recordcount 
				  If rs_dg.eof Then Exit For
				  %>
        <td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%> ��</td>
        <%
				if i mod 5=0 then
				response.write "</tr><tr>"
				end if
				rs_dg.Movenext
				next
				end if
				rs_dg.close
				set rs_dg=nothing
				end if
				%>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="15%" valign="top">&nbsp;��������������</td>
    <td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
	  set tonglei_rs=server.createobject("adodb.recordset")
	  'sql="Select name,sum(sl) as shuliang From sell_jilu group by name"
	  sql="Select name,sum(sl) as shuliang From sell_jilu where yuangong_id="&conn.execute("select id from yuangong where username='"&userid&"'")(0)&" and "&GetSqlCheckDateString("times")&" group by name"
	  tonglei_rs.open sql,conn,1,1
	  if not tonglei_rs.eof then
	  For i=1 to tonglei_rs.recordcount 
	  If tonglei_rs.eof Then Exit For
	  %>
        <td><%=tonglei_rs("name")%>:&nbsp;<%=tonglei_rs("shuliang")%> ��</td>
        <%
	if i mod 5=0 then
	response.write "</tr><tr>"
	end if
	tonglei_rs.Movenext
	next
	end if
	tonglei_rs.close
	set tonglei_rs=nothing
    %>
      </tr>
    </table></td>
  </tr>
</table>
<%end if%></td>
  </tr>
  <tr>
    <td><%
Response.Write("&nbsp;ͶƱ��&nbsp;&nbsp;")
user_id = conn.execute("select id from yuangong where username='"&userid&"'")(0)

score=60
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��60��;&nbsp;&nbsp;"

score=80
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��80��;&nbsp;&nbsp;"

score=100
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��100��;&nbsp;&nbsp;"
%></td>
  </tr>
</table>
<%
case 5
init_key()
set dict_lf_name=Server.CreateObject("Scripting.Dictionary")
set dict_lf_vol=Server.CreateObject("Scripting.Dictionary")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where (hz_name='"&cur_peplename&"' or hz_name2nd='"&cur_peplename&"') and "&GetSqlCheckDateString("lc_hz"),conn,1,1
%>
<div align="center" style="line-height:30px"> ���ջ�ױ��</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ����</td>
    <td align="center">����/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ӱױ��Ʒ</td>
    <td align="center">��Ӱױ�ɿ�</td>
    <td align="center">��Ƿ��</td>
    <td align="center">�����Ŀ</td>
    <td align="center">ǩ�����</td>
  </tr>
  <%
  idlist=""
  do while not rs.eof
  	jixiang_money=jixiang_money+rs("jixiang_money")
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center">&nbsp;
        <% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		idlist = idlist & ", " & rs("id")
	%>    </td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%  
	hq_money=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(hq_money) then hq_money=0
	hq_allmoney=hq_allmoney+hq_money
	response.write hq_money%>
    </span></font></td>
    <td align="center"><%
	dim rs_pzz,rowinfo
	set rs_pzz = conn.execute("select * from fujia2 where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))
	if not (rs_pzz.eof and rs_pzz.bof) then
		do while not rs_pzz.eof
			rowinfo = GetFieldDataBySQL("select yunyong from yunyong where id="&rs_pzz("jixiang"),"str","N/A")&"/"&rs_pzz("sl")&"��/"&rs_pzz("money")&"Ԫ"
			if rs_pzz("userid")<>userid and rs_pzz("userid2")<>userid and rs_pzz("userid3")<>userid then
				response.write rowinfo&"("&GetFieldDataBySQL("select peplename from yuangong where username='"&rs_pzz("userid")&"'","str","N/A")&")"
			else
				response.write "<font color='red'>"&rowinfo&"</font>"
			end if
			response.write "<br>"
			rs_pzz.movenext
		loop
	else
		response.write "&nbsp;"
	end if
	rs_pzz.close
	set rs_pzz = nothing
	%></td>
    <td align="center"><%
	fj2_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and [type]=3 and xiangmu_id="&rs("id"))(0)
	  if isnull(fj2_money) then fj2_money=0
	  money11=money11+fj2_money
	  response.Write fj2_money
	%></td>
    <td align="center"><%
	dim fm
	fm=FinalMoneySum(rs("id"),false)
	if fm>0 then 
		response.write "<font color=red><b>"&fm&"</b></font>"
	else
		response.write fm
	end if%></td>
    <td align="center"><table width="80%" border="0" cellspacing="0" cellpadding="0">
		<%
	  	if rs("yunyong")="" or isnull(rs("yunyong")) then
	  		response.write "<td>��</td>"
	  	else
	  		yyid=split(rs("yunyong"),", ")
			yysl=split(rs("sl"),", ")
			for yy=0 to ubound(yyid)
				set rsflag = conn.execute("select yunyong from yunyong where type3=1 and id="&yyid(yy))
				if not rsflag.eof then
					'lfcount=lfcount+yysl(yy)
					if dict_lf_name(yyid(yy))<>"" then
						dict_lf_vol(yyid(yy))=dict_lf_vol(yyid(yy))+cint(yysl(yy))
					else
						dict_lf_name(yyid(yy))=rsflag("yunyong")
						dict_lf_vol(yyid(yy))=cint(yysl(yy))
					end if
					%>
				<tr>
                <td>&nbsp;<%=rsflag("yunyong")%></td>
                <td>&nbsp;<%=yysl(yy)%>��&nbsp;</td>
              </tr>
				<%	
				end if
				rsflag.close()
				set rsflag=nothing
			next
		end if
			%>
          </table><%'=GetWedVol(rs("id"))
	%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rslf = server.CreateObject("adodb.recordset")
	rslf.open "SELECT hs_signtype.title, hs_signhistory.vol FROM hs_signtype INNER JOIN hs_signhistory ON hs_signtype.ID = hs_signhistory.typeid where hs_signhistory.userid="&cur_userid&" and hs_signhistory.xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("hs_signhistory.idate"),conn,1,1
	do while not rslf.eof
	%>
      <tr>
        <td>&nbsp;<%=rslf("title")%></td>
        <td align="right"><%=rslf("vol")%>&nbsp;</td>
      </tr>
      <%
		rslf.movenext
	loop
	rslf.close
	set rslf=nothing
	%>
    </table></td>
  </tr>
  <%
  rs.movenext
  i=i+1
loop
rs.close
set rs=nothing
  %>
</table>
<%
msidlist = ""
if idlist<>"" then msidlist=mid(idlist,3)
signwedlist = ShowWedSignStats(msidlist, cur_userid)
if signwedlist<>"" then response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td>ǩ�������"&signwedlist&"</td></tr></table>"
call ShowSuitType(idlist)%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="100" valign="top">&nbsp;�����Ŀ�б�</td>
    <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
		i=0
	  if dict_lf_name.Count>0 then
	  	for each idno in dict_lf_name
	  %>
        <td><%=dict_lf_name(idno)%>:&nbsp;<%=dict_lf_vol(idno)%> ��</td>
        <%
			i=i+1
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
      end if
	  
	set dict_lf_name=nothing
	set dict_lf_vol=nothing
    %>
      </tr>
    </table></td>
  </tr>
</table><br>
<%
init_key()

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where hz_name='"&cur_peplename&"' and "&GetSqlCheckDateString("lc_ky"),conn,1,1
'rs.open "select * from shejixiadan where hz_name='"&cur_peplename&"' and "&GetSqlCheckDateString("lc_hz")&" and "&GetSqlCheckDateString("lc_ky"),conn,1,1
%>
<div align="center" style="line-height:30px">����ѡƬ��</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ����</td>
    <td align="center">����/Ƿ��</td>
    <td align="center">��Ӱױ��Ʒ</td>
    <td align="center">��Ӱױ�ɿ�</td>
    <td align="center">��Ƿ��</td>
    <td align="center">����Ƿ��</td>
  </tr>
  <%
  idlist=""
  allsavemoney=0
  allhqmoney=0
  allhqqk=0
  do while not rs.eof
  	jixiang_money=jixiang_money+rs("jixiang_money")
	taoxi_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&" and [type]=1")(0)
	if isnull(taoxi_save) then taoxi_save=0
	
	
	fujia_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&" and [type]=2 and "&GetSqlCheckDateString("times"))(0)
	if isnull(fujia_save) then fujia_save=0
	
	fujia2_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&" and [type]=3 and "&GetSqlCheckDateString("times"))(0)
	if isnull(fujia2_save) then fujia2_save=0
	goumai_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&" and [type]=4 and "&GetSqlCheckDateString("times"))(0)
	if isnull(goumai_save) then goumai_save=0
	'allsavemoney = allsavemoney + taoxi_save + fujia_save + fujia2_save + goumai_save
	
	money1=conn.execute("select jixiang_money from shejixiadan where id="&rs("id"))(0)
	if isnull(money1) then money1=0
	money2=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	if isnull(money2) then money2=0
	allhqmoney = allhqmoney + money2
	money3=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	if isnull(money3) then money3=0
	money4=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	if isnull(money4) then money4=0
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center">&nbsp;
        <% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		idlist = idlist & ", " & rs("id")
	%>    </td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%response.write money2&"/"& GetNonSaveMoney(rs("id"),2)
	  allhqqk = allhqqk + GetNonSaveMoney(rs("id"),2)
	  %>
    </span></font></td>
    <td align="center"><%
	set rs_pzz = conn.execute("select * from fujia2 where xiangmu_id="&rs("id"))
	if not (rs_pzz.eof and rs_pzz.bof) then
		do while not rs_pzz.eof
			rowinfo = GetFieldDataBySQL("select yunyong from yunyong where id="&rs_pzz("jixiang"),"str","N/A")&"/"&rs_pzz("sl")&"��/"&rs_pzz("money")&"Ԫ"
			if rs_pzz("userid")<>userid and rs_pzz("userid2")<>userid and rs_pzz("userid3")<>userid then
				response.write rowinfo&"("&GetFieldDataBySQL("select peplename from yuangong where username='"&rs_pzz("userid")&"'","str","N/A")&")"
			else
				response.write "<font color='red'>"&rowinfo&"</font>"
			end if
			response.write "<br>"
			rs_pzz.movenext
		loop
	else
		response.write "&nbsp;"
	end if
	rs_pzz.close
	set rs_pzz = nothing
	%></td>
    <td align="center"><%
	  money11=money11+fujia2_save
	  response.Write fujia2_save
	%></td>
    <td align="center"><%
	'fm=FinalMoneySum(rs("id"),false)
	fm = GetNonSaveMoney(rs("id"),0)
	if fm>0 then 
		response.write "<font color=red><b>"&fm&"</b></font>"
	else
		response.write fm
	end if%></td>
    <td align="center"><%
	fm = GetNonSaveMoney(rs("id"),2)
	if fm>0 then 
		response.write "<font color=red><b>"&fm&"</b></font>"
	else
		response.write fm
	end if%></td>
  </tr>
  <%
	 fujia_save11=fujia_save11+fujia_save
  rs.movenext
  i=i+1
loop
rs.close

  %>
  <tr>
    <td colspan="9" bgcolor="#EEEEEE">&nbsp;��ϵ��
      <%response.Write int(jixiang_money)%>
      Ԫ&nbsp; ���º����տ
      <%'response.Write formatnumber(allsavemoney,1,0,0,0)
	  'fujia_save11 = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and j.xp_name='"&cur_peplename&"' and "&GetSqlCheckDateString("s.times")&" and "&GetSqlCheckDateString("j.lc_ky"))(0)
	  'if isnull(fujia_save11) then fujia_save11=0
	  hqbk_money = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=2 and hz_name='"&cur_peplename&"' and "&GetSqlCheckDateString("s.times")&" and s.xiangmu_id not in (select id from shejixiadan where "&GetSqlCheckDateString("lc_ky")&")")(0)
	  if isnull(hqbk_money) then hqbk_money  = 0
	  response.Write formatnumber(fujia_save11,1,0,0,0)&" + "& hqbk_money &" (���ڲ���)"%>
Ԫ &nbsp; ѡƬ����Ƿ�<%=formatnumber(allhqqk,1,0,0,0)%>Ԫ &nbsp;&nbsp;��Ƭ������
	  <%
	  if idlist<>"" then 
	  	  tmpidlist = mid(idlist,3)
		  gpsl = conn.execute("select sum(fujia.sl) from fujia inner join yunyong on fujia.jixiang=yunyong.id where yunyong.isgp=1 and fujia.xiangmu_id in ("&tmpidlist&")")(0)
		  if isnull(gpsl) then gpsl=0
		  response.write gpsl & " ��"
	  else
	  	  response.write "0 ��"
	  end if
	  
	  %></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;<%
		set rs_ds1 = server.createobject("adodb.recordset")
		set rs_ds2 = server.createobject("adodb.recordset")
		set rs_ds3 = server.createobject("adodb.recordset")
		
		rs_ds1.open "select distinct s.id from shejixiadan s inner join kehu k on s.kehu_id=k.id where s.hz_name='"&peplename&"' and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		if not (rs_ds1.eof and rs_ds1.bof) then
			ds1_all = rs_ds1.recordcount
		else
			ds1_all = 0
		end if
		rs_ds1.close
		
		rs_ds3.open "select distinct s.id from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where s.hz_name='"&peplename&"' and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		if not (rs_ds3.eof and rs_ds3.bof) then
			ds3_all = rs_ds3.recordcount
		else
			ds3_all = 0
		end if
		rs_ds3.close
		
		ds2_all = 0
		rs_ds2.open "select s.hz_name,f.money from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where s.hz_name='"&peplename&"' and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
		do while not rs_ds2.eof
			ds2_all = ds2_all + rs_ds2("money")
			rs_ds2.movenext
		loop
		rs_ds2.close
		
		ds_count=0		'����
		ds1_count=0		'ѡƬ��¼����
		ds2_count=0		'ѡƬ���Ѻϼ�
		ds3_count=0		'ѡ�����Ѽ�¼����
		set rslost = conn.execute("select * from CustomerLostType order by px")
		do while not rslost.eof
			ds1 = 0
			ds2 = 0
			ds3 = 0
			
			rs_ds1.open "select distinct s.id from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.CustomerLostType="&rslost("id")&" and s.hz_name='"&peplename&"' and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			if not (rs_ds1.eof and rs_ds1.bof) then
				ds1 = rs_ds1.recordcount
			else
				ds1 = 0
			end if
			rs_ds1.close
			
			rs_ds3.open "select distinct s.id from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where k.CustomerLostType="&rslost("id")&" and s.hz_name='"&peplename&"' and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			if not (rs_ds3.eof and rs_ds3.bof) then
				ds3 = rs_ds3.recordcount
			else
				ds3 = 0
			end if
			rs_ds3.close
			
			rs_ds2.open "select s.ky_name,s.ky_name2,f.money from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia f on s.id = f.xiangmu_id where k.CustomerLostType="&rslost("id")&" and s.hz_name='"&peplename&"' and "&GetSqlCheckDateString("f.times")&" and "&GetSqlCheckDateString("s.lc_ky"),conn,1,1
			do while not rs_ds2.eof
				ds2 = ds2 + rs_ds2("money")
				rs_ds2.movenext
			loop
			rs_ds2.close
			
			ds1_count = ds1_count + ds1
			ds2_count = ds2_count + ds2
			ds3_count = ds3_count + ds3
			response.write rslost("title")&"ѡƬ"&ds1&"�� "
			response.write "δ����"& ds1-ds3 &"�� "
			response.write "��"&ds2&"Ԫ ƽ�����"
			if ds1=0 then 
				response.write ".0"
			else
				response.write formatnumber(ds2/ds1,1,0,0,0)
			end if
			response.write " Ԫ&nbsp;&nbsp;&nbsp;"
			ds_count = ds_count + 1
			if ds_count mod 2 = 0 then response.write "<br>&nbsp;"
			rslost.movenext
		loop
		rslost.close
		set rslost = nothing
		
		response.write "����ѡƬ"&ds1_all-ds1_count&"�� "
		response.write "δ����"& (ds1_all-ds3_all)-(ds1_count-ds3_count) &"�� "
		response.write "��"& ds2_all-ds2_count &"Ԫ ƽ�����"
		if (ds1_all-ds3_all)-(ds1_count-ds3_count)=0 then 
			response.write ".0"
		else
			response.write formatnumber((ds2_all-ds2_count)/(ds1_all-ds1_count),1,0,0,0)
		end if
		response.write " Ԫ"
%></td>
  </tr>
</table>
<br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" align="center"><%response.write datearea%>
      &nbsp;������ѡƬ������ϸ��</td>
  </tr>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" style="richness:1px">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">����ϵ��</td>
    <td align="center">��ϵ�ɷ�/(�Ŷ�)</td>
    <td width="16%" align="center">ѡƬ�����ܽ��</td>
    <td align="center">���ڽɷ�/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ƭ���͡�</td>
    <td align="center">��Ƿ��<font color="#FF0000"><span class="style5"></span></font></td>
  </tr>
  <%
  Call init_key()
  rs.open "select * from shejixiadan where hz_name='"&cur_peplename&"' and "&sql_time&" and not ("&GetSqlCheckDateString("lc_hz")&") and id in (select xiangmu_id from save_money where [type]=2 and "&GetSqlCheckDateString("times")&")",conn,1,1
  
  'response.write "select * from shejixiadan where hz_name='"&cur_peplename&"' and "&sql_time&" and not ("&GetSqlCheckDateString("lc_hz")&") and id in (select xiangmu_id from save_money where [type]=2 and "&GetSqlCheckDateString("times")&")"

  'msidlist=","
  do while not rs.eof
  str_sm=""
  
  '�������½ɺ��ڿ�
  hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
  	if isnull(money2) then money2=0

	sm2_money=money2
  
  '�����ܺ���
  hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
  if isnull(hq_money) then hq_money = 0
  
  '�����ܺ��ڽɿ�
  hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
  if isnull(hq_savemoney) then hq_savemoney = 0
  
  	if hq_money=hq_savemoney then
  		ReceivablesMoney = ReceivablesMoney + hq_money
  	end if

  'if hq_money=hq_indate_savemoney then 
  '	RecFujiaMoney = RecFujiaMoney+hq_mymoney
	'AllRecFujiaMoney = AllRecFujiaMoney+hq_money
  'end if
  
  'hq_minesavemoney = conn.execute("select sum(money) from save_money where [type]=2 and userid='"&userid&"' and xiangmu_id="&rs("id"))(0)
  set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
  do while not rshq.eof
  	'if rshq("userid")=userid or rshq("userid2")=userid then
	'  if rshq("userid")<>"" and not isnull(rshq("userid2")) then
	'	hq_mymoney = hq_mymoney + rshq("money")/2
	'  else
	  	hq_mymoney = hq_mymoney + rshq("money")
  	'  end if
	'end if
	rshq.movenext
  loop
  rshq.close
  set rshq=nothing
  if isnull(hq_mymoney) then hq_mymoney = 0
  
  '��Ƿ��
  hq_notsavemoney=hq_notsavemoney+hq_money-hq_savemoney
  
  '�ܺ���
  hq_allmoney=hq_allmoney+hq_money
  
  '�����ܺ��ڽɿ�
  hq_indate_allsavemoney=hq_indate_allsavemoney+hq_indate_savemoney
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center"><% 
		response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		msidlist=msidlist&rs("id")&","
	%>
    </td>
    <td align="center"><%
	 response.Write conn.execute("select lxpeple from kehu where id="&rs("kehu_id"))(0)
	 if count111>1 then response.Write "/<font color=red>�Ŷ�</font>"
	 %></td>
    <td align="center"><% 
		jx_money = rs("jixiang_money")
		response.Write formatnumber(jx_money,1,0,0,0)
	%></td>
    <td align="center"><%money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money1) then money1=0
	response.Write formatnumber(sm1_money,1,0,0,0)
	%></td>
    <td align="center" bgcolor="#ffffff"><%
	hqallmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
	if isnull(hqallmoney) then hqallmoney=0
	response.write Formatnumber(hqallmoney,1,0,0,0)
	%></td>
    <td align="center"><%
	money2=conn.execute("select sum(money) from save_money where type=2 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id"))(0)
	response.Write formatnumber(hq_indate_savemoney,1,0,0,0)
'	if rs("ky_name")<>cur_peplename then
'			response.Write "/"&rs("ky_name")
'	  end if
'	  if rs("ky_name2")<>cur_peplename then
'			response.Write "/"&rs("ky_name2")
'	  end if
	%></td>
    <td align="center" bgcolor="#ffffff"><%if rs("hz_name")<>cur_peplename then
		response.write "0"
	else%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
          <tr>
            <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
            <td>&nbsp;<%=rsdg("all_sl")%></td>
          </tr>
          <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
        </table>
      <%end if%></td>
    <td align="center"><%=FinalMoneySum(rs("id"),False)%></td>
  </tr>
  <%
  money11=money11+sm1_money
  money22=money22+sm2_money
  money33=money33+sm3_money
  money44=money44+sm4_money
 
  rs.movenext
  i=i+1
loop
rs.close()
  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;ѡƬ���� <%=Formatnumber(hq_indate_allsavemoney,1,0,0,0)%> Ԫ&nbsp;&nbsp;&nbsp; &nbsp;�ۼƺ���ѡƬǷ��
      <%
	tmp_fujia_money = conn.execute("select sum(f.money) from fujia f inner join shejixiadan s on f.xiangmu_id=s.id where (s.hz_name='"&cur_peplename&"')")(0)
	tmp_save_money = conn.execute("select sum(m.money) from save_money m inner join shejixiadan s on m.xiangmu_id=s.id where m.type=2 and (s.hz_name='"&cur_peplename&"')")(0)
	if isnull(tmp_fujia_money) then tmp_fujia_money = 0
	if isnull(tmp_save_money) then tmp_save_money = 0
	response.write Formatnumber(tmp_fujia_money-tmp_save_money,1,0,0,0)%>
      Ԫ&nbsp;&nbsp;&nbsp; &nbsp;������� <%=Formatnumber(ReceivablesMoney,1,0,0,0)%> Ԫ</td>
  </tr>
</table>
<div align="center" style="line-height:30px">
  ��黯ױ��</div>
<%init_key()

set dict_lf_name=Server.CreateObject("Scripting.Dictionary")
set dict_lf_vol=Server.CreateObject("Scripting.Dictionary")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where 1=1 and hz_userid='"&userid&"' and"&GetSqlCheckDateString("hz_qm_times"),conn,1,1
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">�ͻ����/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ӱױ��Ʒ</td>
    <td align="center">�ɿ�/Ԫ</td>
    <td align="center">���ױ��Ʒ</td>
    <td align="center">�ɿ�/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ƿ��</td>
	<td align="center">�����Ŀ</td>
    <td align="center">���ͽ��</td>
  </tr>
  <%
  idlist=""
  do while not rs.eof
  	jixiang_money=jixiang_money+rs("jixiang_money")
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center">&nbsp;
        <% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		idlist = idlist & ", " & rs("id")
	%>    </td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%  
	hq_money=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(hq_money) then hq_money=0
	hq_allmoney=hq_allmoney+hq_money
	response.write hq_money%>
    </span></font></td>
    <td align="center"><%
	set rs_pzz = conn.execute("select * from fujia2 where xiangmu_id="&rs("id"))
	if not (rs_pzz.eof and rs_pzz.bof) then
		do while not rs_pzz.eof
			rowinfo = GetFieldDataBySQL("select yunyong from yunyong where id="&rs_pzz("jixiang"),"str","N/A")&"/"&rs_pzz("sl")&"��/"&rs_pzz("money")&"Ԫ"
			if rs_pzz("userid")<>userid and rs_pzz("userid2")<>userid and rs_pzz("userid3")<>userid then
				response.write rowinfo&"("&GetFieldDataBySQL("select peplename from yuangong where username='"&rs_pzz("userid")&"'","str","N/A")&")"
			else
				response.write "<font color='red'>"&rowinfo&"</font>"
			end if
			response.write "<br>"
			rs_pzz.movenext
		loop
	else
		response.write "&nbsp;"
	end if
	rs_pzz.close
	set rs_pzz = nothing
	%></td>
    <td align="center"><%fj2_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and [type]=3 and xiangmu_id="&rs("id"))(0)
	  if isnull(fj2_money) then fj2_money=0
	  money11=money11+fj2_money
	  response.Write fj2_money
	%></td>
    <td align="center"><%
	set rs_jhz = conn.execute("select * from goumai where xiangmu_id="&rs("id"))
	if not (rs_jhz.eof and rs_jhz.bof) then
		do while not rs_jhz.eof
			rowinfo = GetFieldDataBySQL("select yunyong from yunyong where id="&rs_jhz("jixiang"),"str","N/A")&"/"&rs_jhz("sl")&"��/"&rs_jhz("money")&"Ԫ"
			if rs_jhz("userid")<>userid and rs_jhz("userid2")<>userid and rs_jhz("userid3")<>userid then
				response.write rowinfo&"("&GetFieldDataBySQL("select peplename from yuangong where username='"&rs_jhz("userid")&"'","str","N/A")&")"
			else
				response.write "<font color='red'>"&rowinfo&"</font>"
			end if
			response.write "<br>"
			rs_jhz.movenext
		loop
	else
		response.write "&nbsp;"
	end if
	rs_jhz.close
	set rs_jhz = nothing
	%></td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%gm_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and [type]=4 and xiangmu_id="&rs("id"))(0)
	  if isnull(gm_money) then gm_money=0
	  money414=money414+gm_money
	  response.Write gm_money
	%>
    </span></font></td>
    <td align="center"><%
	fm=FinalMoneySum(rs("id"),false)
	if fm>0 then
		response.write "<font color=red><b>"&fm&"</b></font>"
	else
		response.write fm
	end if%></td>
    <td align="center"><table width="80%" border="0" cellspacing="0" cellpadding="0">
		<%
	  	if rs("yunyong")="" or isnull(rs("yunyong")) then
	  		response.write "<td>��</td>"
	  	else
	  		yyid=split(rs("yunyong"),", ")
			yysl=split(rs("sl"),", ")
			for yy=0 to ubound(yyid)
				set rsflag = conn.execute("select yunyong from yunyong where type3=1 and id="&yyid(yy))
				if not rsflag.eof then
					'lfcount=lfcount+yysl(yy)
					if dict_lf_name(yyid(yy))<>"" then
						dict_lf_vol(yyid(yy))=dict_lf_vol(yyid(yy))+cint(yysl(yy))
					else
						dict_lf_name(yyid(yy))=rsflag("yunyong")
						dict_lf_vol(yyid(yy))=cint(yysl(yy))
					end if
					%>
				<tr>
                <td>&nbsp;<%=rsflag("yunyong")%></td>
                <td>&nbsp;<%=yysl(yy)%>��&nbsp;</td>
              </tr>
				<%	
				end if
				rsflag.close()
				set rsflag=nothing
			next
		end if
			%>
          </table><%'=GetWedVol(rs("id"))
	%></td>
    <td align="center"><%
	dim strps
	strps=""
	if rs("jhz_style")="" or isnull(rs("jhz_style")) then
		strps = "&nbsp;"
	else
		if instr(rs("jhz_style"),"1")>0 then
			strps = strps & "<br>�շ�ױ"
		end if
		if instr(rs("jhz_style"),"2")>0 then
			strps = strps & "<br>���ױ"
		end if
		if strps<>"" then strps = mid(strps,5)
	end if
	response.write strps
	%></td>
  </tr>
  <%
  rs.movenext
  i=i+1
loop
  %>
  <tr>
    <td colspan="10" bgcolor="#EEEEEE">&nbsp;��ϵ��
      <%response.Write int(jixiang_money)%>
      Ԫ&nbsp;&nbsp;&nbsp;&nbsp; ѡƬ��
      <%response.Write hq_allmoney%>
      Ԫ </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="100" valign="top">&nbsp;�����Ŀ�б�</td>
    <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
		i=0
	  if dict_lf_name.Count>0 then
	  	for each idno in dict_lf_name
	  %>
        <td><%=dict_lf_name(idno)%>:&nbsp;<%=dict_lf_vol(idno)%> ��</td>
        <%
			i=i+1
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
      end if
	  
	set dict_lf_name=nothing
	set dict_lf_vol=nothing
    %>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" style="padding-top:10px"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="33%" valign="top"><table width="98%" border="1" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="5" align="center" bgcolor="#EEEEEE">���ջ�ױ</td>
          </tr>
          <tr>
            <td align="center">���</td>
            <td align="center">��Ŀ</td>
            <td align="center">����</td>
            <td align="center">���</td>
            <td align="center">���</td>
          </tr>
          <%
		  if idlist<>"" then
		  	sql_id=" and xiangmu_id in ("&mid(idlist,3)&")"
		  end if
'		  set rs5=server.CreateObject("adodb.recordset")
'		sql="select * from yunyong where id in (select jixiang from fujia2 where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id&")"
'		rs5.open sql,conn,1,1
'		i=0
'		pz_consumer_money = 0
'		while not rs5.eof 
'			i=i+1
'			sl12=conn.execute("select sum(sl) from fujia2 where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id&"")(0)
'			fujia2_money=conn.execute("select sum(money) from fujia2 where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id&"")(0)
'			pz_choucheng=rs5("choucheng")*sl12
'			pz_choucheng11=pz_choucheng11+pz_choucheng
'			pz_consumer_money = pz_consumer_money + fujia2_money
		set rspz = conn.execute("select jixiang,sum(sl) as all_sl,sum(lsmoney) as all_money from (select jixiang,sl,iif(not isnull(userid2) and userid2<>'' and not isnull(userid3) and userid3<>'',money/3,iif(not isnull(userid2) and userid2<>'',money/2,money)) as lsmoney from fujia2 where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&GetSqlCheckDateString("times")&") group by jixiang")
		
		i=0
		pz_consumer_money = 0
		pz_choucheng = 0
		
		do while not rspz.eof
			'e_counter = 0
			set yyrs=conn.execute("select yunyong,choucheng from yunyong where id ="&rspz("jixiang"))
			if not yyrs.eof then
				i=i+1
				'if rspz("userid")<>"" and not isnull(rspz("userid")) then e_counter = e_counter + 1
				'if rspz("userid2")<>"" and not isnull(rspz("userid2")) then e_counter = e_counter + 1
				'if rspz("userid3")<>"" and not isnull(rspz("userid3")) then e_counter = e_counter + 1
				pz_consumer_money = pz_consumer_money + rspz("all_money")'/e_counter
				pz_choucheng = pz_choucheng + yyrs("choucheng")*rspz("all_sl")
		%>
          <tr>
            <td align="center"><%=i%></td>
            <td align="center"><%=yyrs("yunyong")%></td>
            <td align="center"><%=rspz("all_sl")%></td>
            <td align="center"><%=rspz("all_money")%></td>
            <td align="center"><%=yyrs("choucheng")*rspz("all_sl")%></td>
          </tr>
          <%
		  	end if
			yyrs.close
			set yyrs=nothing
			rspz.movenext
		loop
		rspz.close
		set rspz=nothing
			'rs5.movenext
'		wend 
'		rs5.close
'		set rs5=nothing
		
		fujia2_save_money= conn.execute("select sum(money) from save_money where xiangmu_id in (select xiangmu_id from fujia2 where "&GetSqlCheckDateString("times")&") and type=3")(0)
		'conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=3")(0)
		if isnull(fujia2_save_money) then fujia2_save_money=0
		
		'pz_consumer_money = conn.execute("select sum(money) from fujia2 where "&GetSqlCheckDateString("times"))(0)
		if isnull(pz_consumer_money) then pz_consumer_money=0
		%>
          <tr>
            <td colspan="5" bgcolor="#fefefe"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;�����տ
                  <%
				hqsk_money = conn.execute("select sum(money) from save_money where [type]=3 and "&GetSqlCheckDateString("times")&" and userid='"&userid&"'")(0)
				if isnull(hqsk_money) then hqsk_money  = 0
				response.write hqsk_money & "Ԫ"
				'hqxf_money = conn.execute("select sum(money) from fujia2 where "&GetSqlCheckDateString("times")&" and userid='"&userid&"'")(0)
				'if isnull(hqxf_money) then hqxf_money  = 0
				response.write "&nbsp;&nbsp;�����ѣ�"&pz_consumer_money & "Ԫ"
				%><br />
				&nbsp;�����������ڱ��²��<%
				  hqbk_money = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=3 and "&GetSqlCheckDateString("s.times")&" and s.userid='"&userid&"' and s.xiangmu_id not in (select id from shejixiadan where "&GetSqlCheckDateString("lc_hz")&")")(0)
				  if isnull(hqbk_money) then hqbk_money  = 0
				  response.write hqbk_money & " Ԫ"
				  %></td>
                <td align="right" valign="top">���<%=pz_choucheng11%>Ԫ &nbsp;&nbsp;</td>
              </tr>
            </table></td>
          </tr>
        </table></td>
        <td width="33%" valign="top"><table width="98%" border="1" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td colspan="5" align="center" bgcolor="#eeeeee">��黯ױ</td>
          </tr>
          <tr>
            <td align="center">���</td>
            <td align="center">��Ŀ</td>
            <td align="center">����</td>
            <td align="center">���</td>
            <td align="center">���</td>
          </tr>
          <%
		set rsjh = conn.execute("select jixiang,sum(sl) as all_sl,sum(lsmoney) as all_money from (select jixiang,sl,iif(not isnull(userid2) and userid2<>'' and not isnull(userid3) and userid3<>'',money/3,iif(not isnull(userid2) and userid2<>'',money/2,money)) as lsmoney from goumai where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&GetSqlCheckDateString("times")&") group by jixiang")
		
		i=0
		jh_consumer_money = 0
		jh_choucheng = 0
		do while not rsjh.eof
			'e_counter = 0
			set yyrs=conn.execute("select yunyong,choucheng from yunyong where id ="&rsjh("jixiang"))
			if not yyrs.eof then
				i=i+1
				'if rsjh("userid")<>"" and not isnull(rsjh("userid")) then e_counter = e_counter + 1
				'if rsjh("userid2")<>"" and not isnull(rsjh("userid2")) then e_counter = e_counter + 1
				'if rsjh("userid3")<>"" and not isnull(rsjh("userid3")) then e_counter = e_counter + 1
				jh_consumer_money = jh_consumer_money + rsjh("all_money")'/e_counter
				jh_choucheng = jh_choucheng + yyrs("choucheng")*rsjh("all_sl")
		%>
          <tr>
            <td align="center"><%=i%></td>
            <td align="center"><%=yyrs("yunyong")%></td>
            <td align="center"><%=rsjh("all_sl")%></td>
            <td align="center"><%=rsjh("all_money")%></td>
            <td align="center"><%=yyrs("choucheng")*rsjh("all_sl")%></td>
          </tr>
          <%
		  	end if
			yyrs.close
			set yyrs=nothing
			rsjh.movenext
		loop
		rsjh.close
		set rsjh=nothing
		
		'goumai_save_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and xiangmu_id in (select xiangmu_id from goumai where (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id&") and type=4")(0)
		
		goumai_save_money= conn.execute("select sum(money) from save_money where xiangmu_id in (select xiangmu_id from goumai where "&GetSqlCheckDateString("times")&") and type=4")(0)
		'jh_consumer_money = conn.execute("select sum(money) from goumai where "&GetSqlCheckDateString("times"))(0)
		
		'conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&sql_id&" and type=4")(0)
		if isnull(jixiang_choucheng) then jixiang_choucheng=0
		if isnull(fujia_choucheng) then fujia_choucheng=0
		if isnull(hz_choucheng11) then hz_choucheng11=0
		if isnull(pz_choucheng11) then pz_choucheng11=0
		if isnull(goumai_save_money) then goumai_save_money=0
		if isnull(jh_consumer_money) then jh_consumer_money=0
		%>
          <tr>
            <td colspan="5" bgcolor="#fefefe"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;�����տ
                  <%
				hqsk_money = conn.execute("select sum(money) from save_money where [type]=4 and "&GetSqlCheckDateString("times")&" and userid='"&userid&"'")(0)
				if isnull(hqsk_money) then hqsk_money  = 0
				response.write hqsk_money & " Ԫ"
				hqxf_money = conn.execute("select sum(money) from goumai where "&GetSqlCheckDateString("times")&" and userid='"&userid&"'")(0)
				if isnull(hqxf_money) then hqxf_money  = 0
				response.write "&nbsp;&nbsp;�����ѣ�"&hqxf_money & "Ԫ"
				%><br />
				&nbsp;�����������ڱ��²��<%
				  hqbk_money = conn.execute("select sum(money) from save_money s inner join shejixiadan j on s.xiangmu_id=j.id where s.type=4 and "&GetSqlCheckDateString("s.times")&" and s.userid='"&userid&"' and s.xiangmu_id not in (select id from shejixiadan where "&GetSqlCheckDateString("lc_hz")&")")(0)
				  if isnull(hqbk_money) then hqbk_money  = 0
				  response.write hqbk_money & " Ԫ"
				  %></td>
                <td align="right" valign="top">���<%=jh_choucheng%>Ԫ &nbsp;&nbsp;</td>
              </tr>
            </table></td>
          </tr>
        </table></td>
        <td width="33%" align="right" valign="top"><table width="98%" border="1" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="5" align="center" bgcolor="#eeeeee">��ɢ����</td>
          </tr>
          <tr>
            <td align="center">���</td>
            <td align="center">��Ŀ</td>
            <td align="center">����</td>
            <td align="center">�ܽ��</td>
            <td align="center">���</td>
          </tr>
          <%set rs5=server.CreateObject("adodb.recordset")
	  	sql="select distinct xiangmu_id From goumai_jilu where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"'"&sql_id&")"
		rs5.open sql,conn,1,1
		i=0
		while not rs5.eof 
			i=i+1
			set rs_cont=server.CreateObject("adodb.recordset")
			rs_cont.open "select sum(sl),sum(money),sum(choucheng) from goumai_jilu where xiangmu_id="&rs5("xiangmu_id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')",conn,1,1
			sl12=rs_cont(0)
			fujia2_money=rs_cont(1)
			ls_choucheng=rs_cont(2)
			all_fujia_money = all_fujia_money + fujia2_money
			rs_cont.close
			set rs_cont = nothing
		%>
          <tr>
            <td align="center"><%=i%></td>
            <td align="center"><%set rsyy = conn.execute("select xiangmu from save_type where id="&rs5("xiangmu_id")&"")
			if not rsyy.eof then
				response.write rsyy(0)
			else
				response.write "&nbsp;"
			end if
			rsyy.close
			set rsyy=nothing%></td>
            <td align="center"><%=sl12%></td>
            <td align="center"><%=fujia2_money%></td>
            <td align="center"><%=ls_choucheng%></td>
          </tr>
          <%
			rs5.movenext
		wend 
		rs5.close
		set rs5=nothing
		%>
          <tr>
            <td colspan="5" bgcolor="#fefefe">&nbsp;&nbsp;�ϼƣ�<%=all_fujia_money%>Ԫ</td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
init_key()
if qj_flag<>"" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where hz_name='"&cur_peplename&"' or hz_userid='"&userid&"' and wc_name<>'' and not isnull(wc_name) and "&GetSqlCheckDateString("lc_wc"),conn,1,1
%>
<div align="center" style="line-height:30px">
  <%response.write datearea%>
  &nbsp; ȡ���б�</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">����/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ӱױ��Ʒ/�ɿ�<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">���ױ��Ʒ/�ɿ�<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td width="25%" align="center">������Ƭ</td>
  </tr>
  <%
  idlist=""
  do while not rs.eof
  	jixiang_money=jixiang_money+rs("jixiang_money")
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center">&nbsp;
        <% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
	idlist = idlist & ", " & rs("id")
	%>
    </td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%  
	hq_money=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(hq_money) then hq_money=0
	hq_allmoney=hq_allmoney+hq_money
	response.write hq_money%>
    </span></font></td>
    <td align="center">
    <%fj2_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and [type]=3 and xiangmu_id="&rs("id"))(0)
	  if isnull(fj2_money) then fj2_money=0
	  money11=money11+fj2_money
	  response.Write fj2_money
	%>    </td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%gm_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and [type]=4 and xiangmu_id="&rs("id"))(0)
	  if isnull(gm_money) then gm_money=0
	  money414=money414+gm_money
	  response.Write gm_money
	%>
    </span></font></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
      <tr>
        <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
        <td width="30%">&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
      </tr>
      <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
    </table></td>
  </tr>
  <%
  rs.movenext
  i=i+1
loop
  %>
  <tr>
    <td colspan="8" bgcolor="#EEEEEE">&nbsp;��ϵ��
      <%response.Write int(jixiang_money)%>
      Ԫ&nbsp;&nbsp;&nbsp;&nbsp; ѡƬ��
      <%response.Write hq_allmoney%>
      Ԫ </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="7%" valign="top">&nbsp;���ڹ�Ƭ��</td>
    <td width="85%"><table width="90%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
	 if idlist<>"" then
	  set rs_dg=server.createobject("adodb.recordset")
	  sql = "select jixiang,sum(sl) as all_sl from fujia where xiangmu_id in ("&mid(idlist,3)&") and jixiang in (select id from yunyong where isgp=1) and "&GetSqlCheckDateString("times")&" group by jixiang"
	  rs_dg.open sql,conn,1,1
	  if not rs_dg.eof then
	  For i=1 to rs_dg.recordcount 
	  If rs_dg.eof Then Exit For
	  %>
        <td><%=conn.execute("select yunyong from yunyong where id="&rs_dg("jixiang"))(0)%>:&nbsp;<%=rs_dg("all_sl")%> ��</td>
        <%
	if i mod 5=0 then
	response.write "</tr><tr>"
	end if
	rs_dg.Movenext
	next
	end if
	rs_dg.close
	set rs_dg=nothing
	end if
    %>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" style="padding-top:10px"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="33%" valign="top"><table width="98%" border="1" cellspacing="0" cellpadding="0">
            <tr>
              <td colspan="5" align="center" bgcolor="#EEEEEE">���ջ�ױ</td>
            </tr>
            <tr>
              <td align="center">���</td>
              <td align="center">��Ŀ</td>
              <td align="center">����</td>
              <td align="center">�ܽ��</td>
              <td align="center">���</td>
            </tr>
            <%
		  if idlist<>"" then
		  	sql_id=" and xiangmu_id in ("&mid(idlist,3)&")"
		  end if
		  set rs5=server.CreateObject("adodb.recordset")
		sql="select * from yunyong where id in (select jixiang from fujia2 where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id&")"
		rs5.open sql,conn,1,1
		i=0
		pz_consumer_money = 0
		while not rs5.eof 
			i=i+1
			sl12=conn.execute("select sum(sl) from fujia2 where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id)(0)
			fujia2_money=conn.execute("select sum(money) from fujia2 where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id)(0)
			pz_choucheng=rs5("choucheng")*sl12
			pz_choucheng11=pz_choucheng11+pz_choucheng
			pz_consumer_money = pz_consumer_money + fujia2_money
		%>
            <tr>
              <td align="center"><%=i%></td>
              <td align="center"><%=rs5("yunyong")%></td>
              <td align="center"><%=sl12%></td>
              <td align="center"><%=fujia2_money%></td>
              <td align="center"><%=pz_choucheng%></td>
            </tr>
            <%
			rs5.movenext
		wend 
		rs5.close
		set rs5=nothing
		
		fujia2_save_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and userid='"&userid&"' and type=3")(0)
		if isnull(fujia2_save_money) then fujia2_save_money=0
		%>
            <tr>
              <td colspan="5" bgcolor="#EEEEEE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>&nbsp;&nbsp;����<%=fujia2_save_money%>Ԫ/δ��<%=pz_consumer_money-fujia2_save_money%>Ԫ </td>
                    <td align="right">���<%=pz_choucheng11%>Ԫ &nbsp;&nbsp;</td>
                  </tr>
              </table></td>
            </tr>
        </table></td>
        <td width="33%" valign="top"><table width="98%" border="1" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td colspan="5" align="center" bgcolor="#EEEEEE">��黯ױ</td>
            </tr>
            <tr>
              <td align="center">���</td>
              <td align="center">��Ŀ</td>
              <td align="center">����</td>
              <td align="center">�ܽ��</td>
              <td align="center">���</td>
            </tr>
            <%set rs5=server.CreateObject("adodb.recordset")
		sql="select * from yunyong where id in (select jixiang from goumai where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id&")"
		
		rs5.open sql,conn,1,1
		i=0
		jh_consumer_money = 0
		jh_choucheng11=0
		while not rs5.eof 
			i=i+1
			sl12=conn.execute("select sum(sl) from goumai where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id)(0)
			fujia2_money=conn.execute("select sum(money) from goumai where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"&sql_id)(0)
			if isnull(sl12) then sl12=0
			if isnull(fujia2_money) then fujia2_money=0
			jh_choucheng=rs5("choucheng")*sl12
			if isnull(jh_choucheng) then jh_choucheng=0
			jh_choucheng11=jh_choucheng11+jh_choucheng
			jh_consumer_money = jh_consumer_money + fujia2_money
		%>
            <tr>
              <td align="center"><%=i%></td>
              <td align="center"><%=rs5("yunyong")%></td>
              <td align="center"><%=sl12%></td>
              <td align="center"><%=fujia2_money%></td>
              <td align="center"><%=jh_choucheng%></td>
            </tr>
            <%
			rs5.movenext
		wend 
		rs5.close
		set rs5=nothing
		
		goumai_save_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&sql_id&" and type=4")(0)
		if isnull(jixiang_choucheng) then jixiang_choucheng=0
		if isnull(fujia_choucheng) then fujia_choucheng=0
		if isnull(hz_choucheng11) then hz_choucheng11=0
		if isnull(pz_choucheng11) then pz_choucheng11=0
		if isnull(goumai_save_money) then goumai_save_money=0
		if isnull(jh_consumer_money) then jh_consumer_money=0
		%>
            <tr>
              <td colspan="5" bgcolor="#EEEEEE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>&nbsp;&nbsp;����<%=goumai_save_money%>Ԫ/δ��<%=jh_consumer_money-goumai_save_money%>Ԫ </td>
                    <td align="right">���<%=jh_choucheng11%>Ԫ &nbsp;&nbsp;</td>
                  </tr>
              </table></td>
            </tr>
        </table></td>
        <td width="33%" align="right" valign="top"><table width="98%" border="1" cellspacing="0" cellpadding="0">
            <tr>
              <td colspan="5" align="center" bgcolor="#EEEEEE">��ɢ����</td>
            </tr>
            <tr>
              <td align="center">���</td>
              <td align="center">��Ŀ</td>
              <td align="center">����</td>
              <td align="center">�ܽ��</td>
              <td align="center">���</td>
            </tr>
            <%set rs5=server.CreateObject("adodb.recordset")
	  	sql="select distinct xiangmu_id From goumai_jilu where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"'"&sql_id&")"
		rs5.open sql,conn,1,1
		i=0
		all_fujia_money=0
		while not rs5.eof 
			i=i+1
			set rs_cont=server.CreateObject("adodb.recordset")
			rs_cont.open "select sum(sl),sum(money),sum(choucheng) from goumai_jilu where xiangmu_id="&rs5("xiangmu_id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')",conn,1,1
			sl12=rs_cont(0)
			fujia2_money=rs_cont(1)
			ls_choucheng=rs_cont(2)
			all_fujia_money = all_fujia_money + fujia2_money
			rs_cont.close
			set rs_cont = nothing
		%>
            <tr>
              <td align="center"><%=i%></td>
              <td align="center"><%set rsyy = conn.execute("select xiangmu from save_type where id="&rs5("xiangmu_id")&"")
			if not rsyy.eof then
				response.write rsyy(0)
			else
				response.write "&nbsp;"
			end if
			rsyy.close
			set rsyy=nothing%></td>
              <td align="center"><%=sl12%></td>
              <td align="center"><%=fujia2_money%></td>
              <td align="center"><%=ls_choucheng%></td>
            </tr>
            <%
			rs5.movenext
		wend 
		rs5.close
		set rs5=nothing
		%>
            <tr>
              <td colspan="5" bgcolor="#EEEEEE">&nbsp;&nbsp;�ϼƣ�<%=all_fujia_money%>Ԫ</td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
  <%end if%>
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;���¹���:
      <%
	  if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
	else
		gongzi=0
	end if
end if
if (fromtime<>"" and not isnull(fromtime)) and (totime<>"" and not isnull(totime)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
	else
		gongzi=0
	end if
end if
response.Write gongzi%>
      Ԫ      &nbsp;��ע:
      <%if beizhu="" or isnull(beizhu) then 
response.Write "��"
else
response.Write beizhu
end if%>
      <br>
      &nbsp;
      <%
	  Call showYxTable()
	  
Response.Write("&nbsp;ͶƱ��&nbsp;&nbsp;")
user_id = conn.execute("select id from yuangong where username='"&userid&"'")(0)

score=60
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��60��;&nbsp;&nbsp;"

score=80
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��80��;&nbsp;&nbsp;"

score=100
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��100��;&nbsp;&nbsp;"

%></td>
  </tr>
</table>
<%
case 14
init_key()
set dict_lf_name=Server.CreateObject("Scripting.Dictionary")
set dict_lf_vol=Server.CreateObject("Scripting.Dictionary")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where (hz_name='"&cur_peplename&"' or hz_name2nd='"&cur_peplename&"' or hz_name2='"&cur_peplename&"') and "&GetSqlCheckDateString("lc_hz"),conn,1,1
%>
<div align="center" style="line-height:30px"> ���ջ�ױ��</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ����</td>
    <td align="center">����/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ӱױ��Ʒ</td>
    <td align="center">��Ӱױ�ɿ�</td>
    <td align="center">��Ƿ��</td>
    <td align="center">�����Ŀ</td>
    <td align="center">ǩ�����</td>
  </tr>
  <%
  idlist=""
  do while not rs.eof
  	jixiang_money=jixiang_money+rs("jixiang_money")
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center">&nbsp;
        <% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		idlist = idlist & ", " & rs("id")
	%>    </td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%  
	hq_money=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(hq_money) then hq_money=0
	hq_allmoney=hq_allmoney+hq_money
	response.write hq_money%>
    </span></font></td>
    <td align="center"><%
	set rs_pzz = conn.execute("select * from fujia2 where xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))
	if not (rs_pzz.eof and rs_pzz.bof) then
		do while not rs_pzz.eof
			rowinfo = GetFieldDataBySQL("select yunyong from yunyong where id="&rs_pzz("jixiang"),"str","N/A")&"/"&rs_pzz("sl")&"��/"&rs_pzz("money")&"Ԫ"
			if rs_pzz("userid")<>userid and rs_pzz("userid2")<>userid and rs_pzz("userid3")<>userid then
				response.write rowinfo&"("&GetFieldDataBySQL("select peplename from yuangong where username='"&rs_pzz("userid")&"'","str","N/A")&")"
			else
				response.write "<font color='red'>"&rowinfo&"</font>"
			end if
			response.write "<br>"
			rs_pzz.movenext
		loop
	else
		response.write "&nbsp;"
	end if
	rs_pzz.close
	set rs_pzz = nothing
	%></td>
    <td align="center"><%
	fj2_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and [type]=3 and xiangmu_id="&rs("id"))(0)
	  if isnull(fj2_money) then fj2_money=0
	  money11=money11+fj2_money
	  response.Write fj2_money
	%></td>
    <td align="center"><%
	fm=FinalMoneySum(rs("id"),false)
	if fm>0 then 
		response.write "<font color=red><b>"&fm&"</b></font>"
	else
		response.write fm
	end if%></td>
    <td align="center"><table width="80%" border="0" cellspacing="0" cellpadding="0">
		<%
	  	if rs("yunyong")="" or isnull(rs("yunyong")) then
	  		response.write "<td>��</td>"
	  	else
	  		yyid=split(rs("yunyong"),", ")
			yysl=split(rs("sl"),", ")
			for yy=0 to ubound(yyid)
				set rsflag = conn.execute("select yunyong from yunyong where type3=1 and id="&yyid(yy))
				if not rsflag.eof then
					'lfcount=lfcount+yysl(yy)
					if dict_lf_name(yyid(yy))<>"" then
						dict_lf_vol(yyid(yy))=dict_lf_vol(yyid(yy))+cint(yysl(yy))
					else
						dict_lf_name(yyid(yy))=rsflag("yunyong")
						dict_lf_vol(yyid(yy))=cint(yysl(yy))
					end if
					%>
				<tr>
                <td>&nbsp;<%=rsflag("yunyong")%></td>
                <td>&nbsp;<%=yysl(yy)%>��&nbsp;</td>
              </tr>
				<%	
				end if
				rsflag.close()
				set rsflag=nothing
			next
		end if
			%>
          </table><%'=GetWedVol(rs("id"))
	%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rslf = server.CreateObject("adodb.recordset")
	rslf.open "SELECT hs_signtype.title, hs_signhistory.vol FROM hs_signtype INNER JOIN hs_signhistory ON hs_signtype.ID = hs_signhistory.typeid where hs_signhistory.userid="&cur_userid&" and hs_signhistory.xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("hs_signhistory.idate"),conn,1,1
	do while not rslf.eof
	%>
      <tr>
        <td>&nbsp;<%=rslf("title")%></td>
        <td align="right"><%=rslf("vol")%>&nbsp;</td>
      </tr>
      <%
		rslf.movenext
	loop
	rslf.close
	set rslf=nothing
	%>
    </table></td>
  </tr>
  <%
  rs.movenext
  i=i+1
loop
rs.close
set rs=nothing
  %>
</table>
<%
msidlist = ""
if idlist<>"" then msidlist=mid(idlist,3)
signwedlist = ShowWedSignStats(msidlist, cur_userid)
if signwedlist<>"" then response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td>ǩ�������"&signwedlist&"</td></tr></table>"
call ShowSuitType(idlist)%>
<%call ShowSuitType(idlist)%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="100" valign="top">&nbsp;�����Ŀ�б�</td>
    <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
		i=0
	  if dict_lf_name.Count>0 then
	  	for each idno in dict_lf_name
	  %>
        <td><%=dict_lf_name(idno)%>:&nbsp;<%=dict_lf_vol(idno)%> ��</td>
        <%
			i=i+1
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
      end if
	  
	set dict_lf_name=nothing
	set dict_lf_vol=nothing
    %>
      </tr>
    </table></td>
  </tr>
</table><!--
<%

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where (hz_name='"&cur_peplename&"' or hz_name2='"&cur_peplename&"') and "&GetSqlCheckDateString("lc_hz"),conn,1,1
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td width="11%" height="19"><div align="left">&nbsp;&nbsp;����</div></td>
    <td width="17%"><div align="center">�ͻ�</div></td>
    <td width="18%"><div align="center">��ϵ�ɿ�/<font color="#FF0000"><span class="style5">Ԫ</span></font></div></td>
    <td width="14%"><div align="center">���ڽɿ�/<font color="#FF0000"><span class="style5">Ԫ</span></font></div></td>
    <td width="12%"><div align="center">��Ӱױ�ɿ�</div></td>
    <td width="11%"><div align="center">���ױ�ɿ�</div></td>
    <td width="17%"><div align="center">�µ�ʱ��</div></td>
  </tr>
  <%do while not rs.eof
  %>
  <tr bgcolor="#FFFFFF">
    <td>
      <div align="left"> &nbsp;
          <% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
	
	%>
    </div></td>
    <td><div align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></div></td>
    <td><div align="center">
<%num=conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=5")(0)
	jixiang_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=1 and xiangmu_id="&rs("id")&"")(0)
	if isnull(jixiang_save) then jixiang_save=0
	response.Write int(jixiang_save)
		%>
    </div></td>
    <td><div align="center"><font color="#FF0000"><span class="style5">
<%
fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and  type=2 and xiangmu_id="&rs("id")&"")(0)
	  if isnull(fujia_save) then fujia_save=0
	  if conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=5 and userid='"&userid&"'")(0)=0 then fujia_save=0
	  response.Write fujia_save%>
    </span></font></div></td>
    <td><div align="center"><span class="style5"> </span><font color="#FF0000"><span class="style5">
	<%fujia2_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=3 and xiangmu_id="&rs("id")&"")(0)
	  if isnull(fujia2_save) then fujia2_save=0
	  response.Write fujia2_save
xfujia2_money=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
		if isnull(xfujia2_money) then xfujia2_money=0
		xxfujia2_money=xxfujia2_money+xfujia2_money
			  
	  
	 %>
    </span></font></div></td>
    <td><div align="center"><font color="#FF0000"><span class="style5">
      <%money4=conn.execute("select sum(money) from save_money where type=4 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money4) then money4=0
	if rs("hz_userid")<>userid then money4=0
	response.Write money4%>
    </span></font></div></td>
    <td>
      <div align="center"><%=datevalue(rs("times"))%></div></td>
  </tr>
  <%
   fujia_save11=fujia_save11+fujia_save
    jixiang_money=jixiang_money+jixiang_save
	fujia2_save11=fujia2_save11+fujia2_save
	money414=money414+money4
	goumai_money=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
		if isnull(goumai_money) then goumai_money=0
	xgoumai_money=xgoumai_money+goumai_money	
  rs.movenext
  i=i+1
loop
  %>
  <tr>
  	<td colspan="7" bgcolor="#EEEEEE">&nbsp;��ϵ��
  	  <%response.Write int(jixiang_money)
	jixiang_choucheng=int(jixiang_money)*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
	%>
Ԫ&nbsp;&nbsp;&nbsp;&nbsp; ѡƬ��
<%response.Write fujia_save11 
	fujia_choucheng=fujia_save11*conn.execute("select choucheng2 from yuangong where username='"&userid&"'")(0)
	%>
Ԫ</td>
  </tr>
</table>-->
<div align="center" style="line-height:30px">
  ��黯ױ��</div>
<%init_key()

set dict_lf_name=Server.CreateObject("Scripting.Dictionary")
set dict_lf_vol=Server.CreateObject("Scripting.Dictionary")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where 1=1 and hz_userid='"&userid&"' and"&GetSqlCheckDateString("hz_qm_times"),conn,1,1
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">����</td>
    <td align="center">�ͻ�</td>
    <td align="center">�ͻ����/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ӱױ��Ʒ</td>
    <td align="center">�ɿ�/Ԫ</td>
    <td align="center">���ױ��Ʒ</td>
    <td align="center">�ɿ�/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">��Ƿ��</td>
	<td align="center">�����Ŀ</td>
    <td align="center">���ͽ��</td>
  </tr>
  <%
  idlist=""
  do while not rs.eof
  	jixiang_money=jixiang_money+rs("jixiang_money")
  %>
  <tr bgcolor="#FFFFFF">
    <td align="center">&nbsp;
        <% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"
		idlist = idlist & ", " & rs("id")
	%>    </td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%  
	hq_money=conn.execute("select sum(money) from fujia where "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(hq_money) then hq_money=0
	hq_allmoney=hq_allmoney+hq_money
	response.write hq_money%>
    </span></font></td>
    <td align="center"><%
	set rs_pzz = conn.execute("select * from fujia2 where xiangmu_id="&rs("id"))
	if not (rs_pzz.eof and rs_pzz.bof) then
		do while not rs_pzz.eof
			rowinfo = GetFieldDataBySQL("select yunyong from yunyong where id="&rs_pzz("jixiang"),"str","N/A")&"/"&rs_pzz("sl")&"��/"&rs_pzz("money")&"Ԫ"
			if rs_pzz("userid")<>userid and rs_pzz("userid2")<>userid and rs_pzz("userid3")<>userid then
				response.write rowinfo&"("&GetFieldDataBySQL("select peplename from yuangong where username='"&rs_pzz("userid")&"'","str","N/A")&")"
			else
				response.write "<font color='red'>"&rowinfo&"</font>"
			end if
			response.write "<br>"
			rs_pzz.movenext
		loop
	else
		response.write "&nbsp;"
	end if
	rs_pzz.close
	set rs_pzz = nothing
	%></td>
    <td align="center"><%fj2_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and [type]=3 and xiangmu_id="&rs("id"))(0)
	  if isnull(fj2_money) then fj2_money=0
	  money11=money11+fj2_money
	  response.Write fj2_money
	%></td>
    <td align="center"><%
	set rs_jhz = conn.execute("select * from goumai where xiangmu_id="&rs("id"))
	if not (rs_jhz.eof and rs_jhz.bof) then
		do while not rs_jhz.eof
			rowinfo = GetFieldDataBySQL("select yunyong from yunyong where id="&rs_jhz("jixiang"),"str","N/A")&"/"&rs_jhz("sl")&"��/"&rs_jhz("money")&"Ԫ"
			if rs_jhz("userid")<>userid and rs_jhz("userid2")<>userid and rs_jhz("userid3")<>userid then
				response.write rowinfo&"("&GetFieldDataBySQL("select peplename from yuangong where username='"&rs_jhz("userid")&"'","str","N/A")&")"
			else
				response.write "<font color='red'>"&rowinfo&"</font>"
			end if
			response.write "<br>"
			rs_jhz.movenext
		loop
	else
		response.write "&nbsp;"
	end if
	rs_jhz.close
	set rs_jhz = nothing
	%></td>
    <td align="center"><font color="#FF0000"><span class="style5">
      <%gm_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and [type]=4 and xiangmu_id="&rs("id"))(0)
	  if isnull(gm_money) then gm_money=0
	  money414=money414+gm_money
	  response.Write gm_money
	%>
    </span></font></td>
    <td align="center"><%
	fm=FinalMoneySum(rs("id"),false)
	if fm>0 then
		response.write "<font color=red><b>"&fm&"</b></font>"
	else
		response.write fm
	end if%></td>
    <td align="center"><table width="80%" border="0" cellspacing="0" cellpadding="0">
		<%
	  	if rs("yunyong")="" or isnull(rs("yunyong")) then
	  		response.write "<td>��</td>"
	  	else
	  		yyid=split(rs("yunyong"),", ")
			yysl=split(rs("sl"),", ")
			for yy=0 to ubound(yyid)
				set rsflag = conn.execute("select yunyong from yunyong where type3=1 and id="&yyid(yy))
				if not rsflag.eof then
					'lfcount=lfcount+yysl(yy)
					if dict_lf_name(yyid(yy))<>"" then
						dict_lf_vol(yyid(yy))=dict_lf_vol(yyid(yy))+cint(yysl(yy))
					else
						dict_lf_name(yyid(yy))=rsflag("yunyong")
						dict_lf_vol(yyid(yy))=cint(yysl(yy))
					end if
					%>
				<tr>
                <td>&nbsp;<%=rsflag("yunyong")%></td>
                <td>&nbsp;<%=yysl(yy)%>��&nbsp;</td>
              </tr>
				<%	
				end if
				rsflag.close()
				set rsflag=nothing
			next
		end if
			%>
          </table><%'=GetWedVol(rs("id"))
	%></td>
    <td align="center"><%
	strps = ""
	if rs("jhz_style")="" or isnull(rs("jhz_style")) then
		strps = "&nbsp;"
	else
		if instr(rs("jhz_style"),"1")>0 then
			strps = strps & "<br>�շ�ױ"
		end if
		if instr(rs("jhz_style"),"2")>0 then
			strps = strps & "<br>���ױ"
		end if
		if strps<>"" then strps = mid(strps,5)
	end if
	response.write strps
	%></td>
  </tr>
  <%
  rs.movenext
  i=i+1
loop
  %>
  <tr>
    <td colspan="10" bgcolor="#EEEEEE">&nbsp;��ϵ��
      <%response.Write int(jixiang_money)%>
      Ԫ&nbsp;&nbsp;&nbsp;&nbsp; ѡƬ��
      <%response.Write hq_allmoney%>
      Ԫ </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="100" valign="top">&nbsp;�����Ŀ�б�</td>
    <td><table width="85%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
		i=0
	  if dict_lf_name.Count>0 then
	  	for each idno in dict_lf_name
	  %>
        <td><%=dict_lf_name(idno)%>:&nbsp;<%=dict_lf_vol(idno)%> ��</td>
        <%
			i=i+1
			if i mod 4=0 then
				response.write "</tr><tr>"
			end if
		next
      end if
	  
	set dict_lf_name=nothing
	set dict_lf_vol=nothing
    %>
      </tr>
    </table></td>
  </tr>
</table>
<%'------------------------------------------------------------------
Call showYxTable()
Call showSubTable()

case 11
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where id in (select xiangmu_id from save_money where "&GetSqlCheckDateString("times")&") and (id in (select xiangmu_id from xiadan where userid2='"&userid&"'))",conn,1,1
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF" align="center">
    <td height="19">����</td>
    <td>�ͻ�</td>
    <td>��ϵ�ɿ�/<font color="#FF0000">Ԫ</font></td>
    <td>���ڽɿ�/<font color="#FF0000">Ԫ</font></td>
    <td>��Ӱױ�ɿ�/<font color="#FF0000">Ԫ</font></td>
    <td>���ױ�ɿ�/<font color="#FF0000">Ԫ</font></td>
    <td align="center">�������</td>
  </tr>
  <%do while not rs.eof
  %>
  <tr bgcolor="#FFFFFF" align="center">
    <td><% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"%>
    </td>
    <td><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%></td>
    <td>
  <%num=conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=5")(0)
	jixiang_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=1 and xiangmu_id="&rs("id")&"")(0)
	if isnull(jixiang_save) then jixiang_save=0
	response.Write int(jixiang_save)
		%>
    </td>
    <td><font color="#FF0000"><span class="style5">
  <%
fujia_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and  type=2 and xiangmu_id="&rs("id")&"")(0)
	  if isnull(fujia_save) then fujia_save=0
	  if conn.execute("select count(*) from xiadan where xiangmu_id="&rs("id")&" and type=5 and userid='"&userid&"'")(0)=0 then fujia_save=0
	  response.Write fujia_save%>
    </span></font></td>
    <td><span class="style5"> </span><font color="#FF0000"><span class="style5">
      <%fujia2_save=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and type=3 and xiangmu_id="&rs("id")&"")(0)
	  if isnull(fujia2_save) then fujia2_save=0
	  response.Write fujia2_save
xfujia2_money=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
		if isnull(xfujia2_money) then xfujia2_money=0
		xxfujia2_money=xxfujia2_money+xfujia2_money
	 %>
    </span></font></td>
    <td><font color="#FF0000"><span class="style5">
      <%money4=conn.execute("select sum(money) from save_money where type=4 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	if isnull(money4) then money4=0
	if rs("hz_userid")<>userid then money4=0
	response.Write money4%>
    </span></font></td>
    <td align="center"><%=GetWedVol(rs("id"))%></td>
  </tr>
  <%
   fujia_save11=fujia_save11+fujia_save
    jixiang_money=jixiang_money+jixiang_save
	fujia2_save11=fujia2_save11+fujia2_save
	money414=money414+money4
	goumai_money=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
		if isnull(goumai_money) then goumai_money=0
	xgoumai_money=xgoumai_money+goumai_money	
  rs.movenext
  i=i+1
loop
  %>
  <tr>
  	<td colspan="7" bgcolor="#EEEEEE">&nbsp;��ϵ��
  	  <%response.Write int(jixiang_money)
	jixiang_choucheng=int(jixiang_money)*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
	%>
Ԫ&nbsp;&nbsp;&nbsp;&nbsp; ѡƬ��
<%response.Write fujia_save11 
	fujia_choucheng=fujia_save11*conn.execute("select choucheng2 from yuangong where username='"&userid&"'")(0)
	%>
Ԫ	</td>
  </tr>
</table>
<%
Call showYxTable()
Call showSubTable()
'------------------------------------------------------------------
case 12
set rs6=server.CreateObject("adodb.recordset")
rs6.open "select * from shejixiadan where xp_name='"&cur_peplename&"' and "&GetSqlCheckDateString("lc_xp"),conn,1,1
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19" align="center">&nbsp;&nbsp;����</td>
    <td align="center">�ͻ�</td>
    <td align="center">��Ƭ����</td>
    <td align="center">��Ƭ���/<font color="#FF0000"><span class="style5">Ԫ</span></font></td>
    <td align="center">ǩ�����</td>
  </tr>
  <%do while not rs6.eof%>
  <tr bgcolor="#FFFFFF">
    <td align="center">&nbsp;<% response.Write rs6("id")%></td>
    <td align="center"><%=conn.execute("select lxpeple from kehu where id="&rs6("kehu_id")&"")(0)%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rsdg = conn.execute("select jixiang,sum(sl) as all_sl,sum(money) as all_money from fujia where xiangmu_id="&rs6("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1) group by jixiang")
	do while not rsdg.eof
	%>
      <tr>
        <td>&nbsp;<%=conn.execute("select yunyong from yunyong where id="&rsdg("jixiang"))(0)%></td>
        <td>&nbsp;<%=rsdg("all_sl")%>��&nbsp;</td>
        <td>&nbsp;<%=rsdg("all_money")%>Ԫ&nbsp;</td>
      </tr>
      <%
		rsdg.movenext
	loop
	rsdg.close
	set rsdg=nothing
	%>
    </table></td>
    <td align="center"><%
	  dgmoney=conn.execute("select sum(money) from fujia where xiangmu_id="&rs6("id")&" and "&GetSqlCheckDateString("times")&" and jixiang in (select id from yunyong where isgp=1)")(0)
	  if isnull(dgmoney) then dgmoney=0
	  response.write dgmoney
	money13=conn.execute("select sum(dj*sl) from sell_jilu where "&GetSqlCheckDateString("times")&"")(0)
	if isnull(money13) then money13=0
	money13=formatnumber(money13,1,0,0,0)
	%></td>
    <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	set rslf = server.CreateObject("adodb.recordset")
	rslf.open "SELECT hs_signtype.title, hs_signhistory.vol FROM hs_signtype INNER JOIN hs_signhistory ON hs_signtype.ID = hs_signhistory.typeid where hs_signhistory.userid="&cur_userid&" and hs_signhistory.xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("hs_signhistory.idate"),conn,1,1
	do while not rslf.eof
	%>
      <tr>
        <td>&nbsp;<%=rslf("title")%></td>
        <td align="right"><%=rslf("vol")%>&nbsp;</td>
      </tr>
      <%
		rslf.movenext
	loop
	rslf.close
	set rslf=nothing
	%>
    </table></td>
  </tr>
  <%
    jixiang_money=jixiang_money+jixiang_save
	money113=money113+money13
  rs6.movenext
  i=i+1
loop

  %>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;��ϵ�ܽ�
      <%response.Write int(jixiang_money)
	jixiang_choucheng=int(jixiang_money)*conn.execute("select choucheng1 from yuangong where username='"&userid&"'")(0)
	%>
    Ԫ&nbsp; &nbsp;�����ܽ�
    <%response.Write fujia_save11
	fujia_choucheng=fujia_save11*int(jixiang_money)*conn.execute("select choucheng2 from yuangong where username='"&userid&"'")(0)
	%>Ԫ&nbsp;&nbsp;&nbsp;�����ܽ��:<%response.Write money113
	daogou_choucheng=money113*conn.execute("select choucheng5 from yuangong where username='"&userid&"'")(0)
  if isnull(jixiang_choucheng) then jixiang_choucheng=0
  if isnull(fujia_choucheng) then fujia_choucheng=0
  if isnull(daogou_choucheng) then  daogou_choucheng=0
	%>Ԫ&nbsp;&nbsp;&nbsp;���¹���Ӱ:
      <%num11=conn.execute("select count(*) from shejixiadan where "&GetSqlCheckDateString("lc_xp")&" and xp_name='"&conn.execute("select peplename from yuangong where username='"&userid&"'")(0)&"'")(0)
	if isnull(num11) then num11=0
	num12=conn.execute("select count(*) from shejixiadan where "&GetSqlCheckDateString("lc_cp")&" and cp_name='"&conn.execute("select peplename from yuangong where username='"&userid&"'")(0)&"'")(0)
	if isnull(num12) then num12=0
	response.Write num11
	%>
      ��&nbsp;���¹���ɫ:<%=num12%>
	��<br>
	&nbsp;���¹���:
    <%if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
	else
		gongzi=0
	end if
end if
if (fromtime<>"" and not isnull(fromtime)) and (totime<>"" and not isnull(totime)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
	else
		gongzi=0
	end if
end if
response.Write gongzi%>
Ԫ&nbsp;&nbsp;��ע:
<%
if beizhu="" or isnull(beizhu) then 
response.Write "��"
else
response.Write beizhu
end if
%>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="15%" valign="top">&nbsp;��������������</td>
	<td width="85%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
      <tr>
	 <%
	  set tonglei_rs=server.createobject("adodb.recordset")
	  sql="Select name,sum(sl) as shuliang From sell_jilu group by name"
	  tonglei_rs.open sql,conn,1,1
	  if not tonglei_rs.eof then
	  For i=1 to tonglei_rs.recordcount 
	  If tonglei_rs.eof Then Exit For
	  %>
    <td><%=tonglei_rs("name")%>:&nbsp;<%=tonglei_rs("shuliang")%> ��</td>
    <%
	if i mod 5=0 then
	response.write "</tr><tr>"
	end if
	tonglei_rs.Movenext
	next
	end if
	tonglei_rs.close
	set tonglei_rs=nothing
    %>
      </tr>
    </table></td>
    </tr>
</table>
<%
Call showYxTable()
Response.Write("&nbsp;ͶƱ��&nbsp;&nbsp;")
user_id = conn.execute("select id from yuangong where username='"&userid&"'")(0)

score=60
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��60��;&nbsp;&nbsp;"

score=80
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��80��;&nbsp;&nbsp;"

score=100
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��100��;&nbsp;&nbsp;"
%></td>
  </tr>
</table>

<%end select
sub showQujianTable(lv)
if qj_flag<>"" then
%>
<div align="center" style="line-height:30px">
  <%response.write datearea%>
&nbsp; ȡ���б�</div>
<%
set rs=server.CreateObject("adodb.recordset")
chk_peplename = conn.execute("select peplename from yuangong where username='"&userid&"'")(0)
select case lv
	case 1
		chk_sql = "(userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"
	case 2
		chk_sql = "sj_name='"&chk_peplename&"'"
	case 4
		chk_sql = "(cp_name='"&chk_peplename&"' or cp_name2='"&chk_peplename&"' or cp_name3='"&chk_peplename&"' or cp_name4='"&chk_peplename&"' or cp_name5='"&chk_peplename&"')"
	case 5
		chk_sql = "(hz_name='"&chk_peplename&"' or hz_userid='"&userid&"')"
end select
rs.open "select * from shejixiadan where "&chk_sql&" and "&GetSqlCheckDateString("lc_wc"),conn,1,1
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF" align="center">
    <td height="19">����</td>
    <td>�ͻ�/���� </td>
    <td>��ϵ</td>
    <td>����</td>
    <td>�Ŵ�</td>
    <td align="center" valign="middle">���ȡ��</td>
  </tr>
  <%
   banmianll=0
   fangdall=0
   idlist=""
  do while not rs.eof
  	idlist=idlist&","&rs("id")
    save_money=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id")&"")(0)
	if isnull(save_money) then save_money=0
	fujia1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia1) then fujia1=0
	fujia2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id")&"")(0)
	if isnull(fujia2) then fujia2=0
	goumai=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id")&"")(0)
	if isnull(goumai) then goumai=0
	jixiang_money=rs("jixiang_money")
	money111=fujia1+fujia2+jixiang_money-save_money
	
	banmian=rs("banmian")
	if isnull(banmian) then banmian=0
	 fangda=rs("fangda")
	if isnull(fangda) then fangda=0
	
	 %>
  <tr bgcolor="#FFFFFF" align="center">
    <td><% response.write "<a href='javascript:' onClick=""javascript:openkswin('kehu_mianban.asp?id="&rs("id")&"',450,500);"">"&rs("id")&"</a>"%></td>
    <td><%=conn.execute("select lxpeple from kehu where id="&rs("kehu_id")&"")(0)%>/
    <%if money111>0 then 
	response.Write "δ����"
	else
	response.Write "�ѽ���"
	end if
	%></td>
    <td><%=conn.execute("select jixiang from jixiang where id="&rs("jixiang")&"")(0)%></td>

    <td><%=rs("banmian")%>��</td>
    <td><%=rs("fangda")%>��</td>
    <td><%if not isnull(rs("lc_wc")) then
		response.write datevalue(rs("lc_wc"))
	else
		response.write "&nbsp;"
	end if%></td>
  </tr>
  <%
 ' choucheng11=choucheng11+choucheng
   banmianll=banmianll+banmian
  fangdall=fangdall+fangda
 
  jixiang_money=jixiang_money+rs("jixiang_money")
  rs.movenext
  i=i+1
loop
  %>
</table>
<br>
<%
end if
end sub
sub showYxTable()
	dim rsyx
	set rsyx = server.CreateObject("adodb.recordset")
	rsyx.open "select * from shejixiadan where userid3='"&userid&"' and "&GetSqlCheckDateString("times"),conn,1,1
	if not rsyx.eof then
%>
<div align="center" style="line-height:30px">
  <%response.write datearea%>
  &nbsp; Ӫ���б�</div>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#99FFFF">
    <td height="19">&nbsp;����</td>
    <td>&nbsp;����</td>
    <td>&nbsp;��ϵ���</td>
    <td>&nbsp;��ϵ�ɿ�</td>
    <td>&nbsp;ѡƬ����</td>
    <td>&nbsp;���ڽɿ�</td>
  </tr>
  <%
  do while not rsyx.eof
	
	 %>
  <tr bgcolor="#FFFFFF">
    <td>&nbsp;<%=rsyx("id")%></td>
    <td>&nbsp;<%
	msname=""
	set rsms = conn.execute("select peplename from yuangong where username='"&rsyx("userid")&"'")
	if not (rsms.eof and rsms.bof) then
		msname = rsms("peplename")
	end if
	rsms.close
	if rsyx("userid2")<>"" and not isnull(rsyx("userid2")) then
		set rsms = conn.execute("select peplename from yuangong where username='"&rsyx("userid2")&"'")
		if not (rsms.eof and rsms.bof) then
			if msname<>"" then msname = msname & "/" &rsms("peplename")
		else
			msname = rsms("peplename")
		end if
		rsms.close
	end if
	set rsms = nothing
	response.write msname
	%></td>
    <td>&nbsp;<%=rsyx("jixiang_money")%></td>
    <td>&nbsp;<%
	taoxi_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rsyx("id")&" and [type]=1")(0)
	if isnull(taoxi_save) then taoxi_save=0
	response.write taoxi_save
	%></td>
    <td>&nbsp;<%
	money2=conn.execute("select sum(money) from fujia where xiangmu_id="&rsyx("id"))(0)
	if isnull(money2) then money2=0
	response.write money2
	%></td>
    <td>&nbsp;<%
	fujia_save=conn.execute("select sum(money) from save_money where xiangmu_id="&rsyx("id")&" and [type]=2")(0)
	if isnull(fujia_save) then fujia_save=0
	response.write fujia_save
	%></td>
  </tr>
  <%

 		rsyx.movenext
	loop
  %>
</table>
<%
end if
rsyx.close
set rsyx=nothing
end sub

sub showSubTable()
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" style="padding-top:10px"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="33%" valign="top"><table width="98%" border="1" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="5" align="center" bgcolor="#EEEEEE">���ջ�ױ</td>
          </tr>
          <tr>
            <td align="center">���</td>
            <td align="center">��Ŀ</td>
            <td align="center">����</td>
            <td align="center">�ܽ��</td>
            <td align="center">���</td>
          </tr>
          <%set rs5=server.CreateObject("adodb.recordset")
		sql="select * from yunyong where id in (select jixiang from fujia2 where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"'))"
		rs5.open sql,conn,1,1
		i=0
		pz_consumer_money = 0
		while not rs5.eof 
			i=i+1
			sl12=conn.execute("select sum(sl) from fujia2 where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')")(0)
			fujia2_money=conn.execute("select sum(money) from fujia2 where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')")(0)
			pz_choucheng=rs5("choucheng")*sl12
			pz_choucheng11=pz_choucheng11+pz_choucheng
			pz_consumer_money = pz_consumer_money + fujia2_money
		%>
          <tr>
            <td align="center"><%=i%></td>
            <td align="center"><%=rs5("yunyong")%></td>
            <td align="center"><%=sl12%></td>
            <td align="center"><%=fujia2_money%></td>
            <td align="center"><%=pz_choucheng%></td>
          </tr>
          <%
			rs5.movenext
		wend 
		rs5.close
		set rs5=nothing
		
		'fujia2_save_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and  xiangmu_id in (select xiangmu_id from xiadan where userid='"&userid&"') and type=3")(0)
		'if isnull(fujia2_save_money) then fujia2_save_money=0
		
		hqsk_money = conn.execute("select sum(money) from save_money where [type]=3 and "&GetSqlCheckDateString("times")&" and userid='"&userid&"'")(0)
		if isnull(hqsk_money) then hqsk_money  = 0
		
		'hqxf_money = conn.execute("select sum(money) from fujia2 where "&GetSqlCheckDateString("times")&" and userid='"&userid&"'")(0)
		'if isnull(hqxf_money) then hqxf_money  = 0
		%>
          <tr>
            <td colspan="5" bgcolor="#EEEEEE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>&nbsp;�����տ<%=hqsk_money%>Ԫ&nbsp;&nbsp;�����ѣ�<%=pz_consumer_money%>Ԫ </td>
                  <td align="right">���<%=pz_choucheng11%>Ԫ &nbsp;&nbsp;</td>
                </tr>
            </table></td>
          </tr>
        </table></td>
        <td width="33%" valign="top"><table width="98%" border="1" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td colspan="5" align="center" bgcolor="#EEEEEE">��黯ױ</td>
          </tr>
          <tr>
            <td align="center">���</td>
            <td align="center">��Ŀ</td>
            <td align="center">����</td>
            <td align="center">�ܽ��</td>
            <td align="center">���</td>
          </tr>
          <%set rs5=server.CreateObject("adodb.recordset")
		sql="select * from yunyong where id in (select jixiang from goumai where "&GetSqlCheckDateString("times")&" and userid='"&userid&"')"
		
		rs5.open sql,conn,1,1
		i=0
		jh_consumer_money = 0
		jh_choucheng11=0
		while not rs5.eof 
			i=i+1
			sl12=conn.execute("select sum(sl) from goumai where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')")(0)
			fujia2_money=conn.execute("select sum(money) from goumai where jixiang="&rs5("id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')")(0)
			if isnull(sl12) then sl12=0
			if isnull(fujia2_money) then fujia2_money=0
			jh_choucheng=rs5("choucheng")*sl12
			if isnull(jh_choucheng) then jh_choucheng=0
			jh_choucheng11=jh_choucheng11+jh_choucheng
			jh_consumer_money = jh_consumer_money + fujia2_money
		%>
          <tr>
            <td align="center"><%=i%></td>
            <td align="center"><%=rs5("yunyong")%></td>
            <td align="center"><%=sl12%></td>
            <td align="center"><%=fujia2_money%></td>
            <td align="center"><%=jh_choucheng%></td>
          </tr>
          <%
			rs5.movenext
		wend 
		rs5.close
		set rs5=nothing
		
		goumai_save_money=conn.execute("select sum(money) from save_money where "&GetSqlCheckDateString("times")&" and  xiangmu_id in (select xiangmu_id from xiadan where userid='"&userid&"') and type=4")(0)
		if isnull(jixiang_choucheng) then jixiang_choucheng=0
		if isnull(fujia_choucheng) then fujia_choucheng=0
		if isnull(hz_choucheng11) then hz_choucheng11=0
		if isnull(pz_choucheng11) then pz_choucheng11=0
		if isnull(goumai_save_money) then goumai_save_money=0
		if isnull(jh_consumer_money) then jh_consumer_money=0
		
		
		hqsk_money = conn.execute("select sum(money) from save_money where [type]=4 and "&GetSqlCheckDateString("times")&" and userid='"&userid&"'")(0)
		if isnull(hqsk_money) then hqsk_money  = 0

		hqxf_money = conn.execute("select sum(money) from goumai where "&GetSqlCheckDateString("times")&" and userid='"&userid&"'")(0)
		if isnull(hqxf_money) then hqxf_money  = 0
		%>
          <tr>
            <td colspan="5" bgcolor="#EEEEEE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>&nbsp;�����տ<%=hqsk_money%>Ԫ&nbsp;&nbsp;�����ѣ�<%=hqxf_money%>Ԫ </td>
                  <td align="right">���<%=jh_choucheng11%>Ԫ &nbsp;&nbsp;</td>
                </tr>
            </table></td>
          </tr>
        </table></td>
        <td width="33%" align="right" valign="top"><table width="98%" border="1" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="5" align="center" bgcolor="#EEEEEE">��ɢ����</td>
          </tr>
          <tr>
            <td align="center">���</td>
            <td align="center">��Ŀ</td>
            <td align="center">����</td>
            <td align="center">�ܽ��</td>
            <td align="center">���</td>
          </tr>
          <%set rs5=server.CreateObject("adodb.recordset")
	  	sql="select distinct xiangmu_id From goumai_jilu where "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')"
		rs5.open sql,conn,1,1
		i=0
		while not rs5.eof 
			i=i+1
			set rs_cont=server.CreateObject("adodb.recordset")
			rs_cont.open "select sum(sl),sum(money),sum(choucheng) from goumai_jilu where xiangmu_id="&rs5("xiangmu_id")&" and "&GetSqlCheckDateString("times")&" and (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"')",conn,1,1
			sl12=rs_cont(0)
			fujia2_money=rs_cont(1)
			ls_choucheng=rs_cont(2)
			all_fujia_money = all_fujia_money + fujia2_money
			rs_cont.close
			set rs_cont = nothing
		%>
          <tr>
            <td align="center"><%=i%></td>
            <td align="center"><%
			dim rsls
			set rsls = conn.execute("select xiangmu from save_type where id="&rs5("xiangmu_id")&"")
			if not (rsls.eof and rsls.bof) then
				response.write rsls("xiangmu")
			else
				response.write "&nbsp;"
			end if
			rsls.close
			set rsls = nothing%></td>
            <td align="center"><%=sl12%></td>
            <td align="center"><%=fujia2_money%></td>
            <td align="center"><%=ls_choucheng%></td>
          </tr>
          <%
			rs5.movenext
		wend 
		rs5.close
		set rs5=nothing
		%>
          <tr>
            <td colspan="5" bgcolor="#EEEEEE">&nbsp;�ϼƣ�<%=all_fujia_money%>Ԫ</td>
          </tr>
        </table></td>
      </tr>
    </table>    </td>
  </tr>
  <tr>
    <td>&nbsp;���¹���:
      <%
	  if (yeard<>"" and not isnull(yeard)) and (monthd<>"" and not isnull(monthd)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&yeard&" and month="&monthd&"")(0)
	else
		gongzi=0
	end if
end if
if (fromtime<>"" and not isnull(fromtime)) and (totime<>"" and not isnull(totime)) then
	if conn.execute("select count(*) from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)>0 then
		gongzi=conn.execute("select money from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
		beizhu=conn.execute("select beizhu from gongzi where userid='"&userid&"' and year="&year(fromtime)&" and month="&month(fromtime))(0)
	else
		gongzi=0
	end if
end if
response.Write gongzi%>
      Ԫ      &nbsp;��ע:
      <%if beizhu="" or isnull(beizhu) then 
response.Write "��"
else
response.Write beizhu
end if%>
      <br>
      <%
	 if typed=1 then 
	 	response.write "<br>"
	   call init_key()
	   set rs=server.CreateObject("adodb.recordset")
	   sql="select * from shejixiadan where ((userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&GetSqlCheckDateString("times")&") or (((userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and "&GetSqlCheckDateString("times")&") and "&GetSqlCheckDateString("lc_cp")&") or (userid='"&userid&"' or userid2='"&userid&"' or userid3='"&userid&"') and (not ("&GetSqlCheckDateString("times")&")) and id in (select xiangmu_id from save_money where "&GetSqlCheckDateString("times")&") or ((ky_name='"&peplename&"' or ky_name2='"&peplename&"') And "&GetSqlCheckDateString("lc_ky")&" and "&GetSqlCheckDateString("lc_cp")&")"
	   rs.open sql,conn,1,1
	     do while not rs.eof
	  str_sm=""
	  if not isnull(rs("userid3")) and rs("userid3")<>"" then 
		count111=3
		elseif not isnull(rs("userid2")) and rs("userid2")<>"" then
		count111=2
		else
		count111=1
		end if
		
		'�������½���ϵ��
  		jx_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=1 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  		if isnull(jx_indate_savemoney) then jx_indate_savemoney=0
	  '��������ϵ
	  jx_money = rs("jixiang_money")
	  
	  '��������ϵ�ɿ�
	  jx_savemoney = conn.execute("select sum(money) from save_money where [type]=1 and xiangmu_id="&rs("id"))(0)
	  
	  if jx_indate_savemoney>0 and jx_money=jx_savemoney then
  		jx_mymoney = jx_mymoney + rs("jixiang_money")/count111
	  end if
		
		
		'�������½ɺ��ڿ�
  		hq_indate_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id")&" and "&GetSqlCheckDateString("times"))(0)
  		if isnull(hq_indate_savemoney) then hq_indate_savemoney=0
	  '�����ܺ���
	  hq_money = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	  if isnull(hq_money) then hq_money = 0
	  
	  '�����ܺ��ڽɿ�
	  hq_savemoney = conn.execute("select sum(money) from save_money where [type]=2 and xiangmu_id="&rs("id"))(0)
	  
	  if hq_indate_savemoney>0 and hq_money=hq_savemoney then
  
		  set rshq = conn.execute("select * from fujia where xiangmu_id="&rs("id"))
		  do while not rshq.eof
			if rshq("userid")=userid or rshq("userid2")=userid then
			  if rshq("userid")<>"" and not isnull(rshq("userid2")) then
				hq_mymoney = hq_mymoney + rshq("money")/2
			  else
				hq_mymoney = hq_mymoney + rshq("money")
			  end if
			end if
			rshq.movenext
		  loop
		  rshq.close
		  set rshq=nothing
	  end if
	  
	  
	  money1=conn.execute("select sum(money) from save_money where type=1 and "&GetSqlCheckDateString("times")&" and xiangmu_id="&rs("id")&"")(0)
	  if isnull(money1) then money1=0
	  
		if rs("userid")=userid or rs("userid2")=userid or rs("userid3")=userid then
			money00=money00+money1/count111
		  end if
	  	rs.movenext
	  loop
	   rs.close
	   set rs=nothing
	   response.write "&nbsp;�ܽ�����ϵ��: "&formatnumber(jx_mymoney,1,0,0,0)&"Ԫ&nbsp;&nbsp;&nbsp; �ܽ�����ڿ�: "&formatnumber(hq_mymoney,1,0,0,0)&" Ԫ"
	   end if%>
      <br>
<%
Response.Write("&nbsp;ͶƱ��&nbsp;&nbsp;")
user_id = conn.execute("select id from yuangong where username='"&userid&"'")(0)

score=60
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��60��;&nbsp;&nbsp;"

score=80
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��80��;&nbsp;&nbsp;"

score=100
Response.Write(Conn.Execute("Select count(*) From Vote Where "&GetSqlCheckDateString("idate")&" and ((ms_user1="&user_id&" and ms_score1="&score&") or (ms_user2="&user_id&" and ms_score2="&score&") or (ms_user3="&user_id&" and ms_score3="&score&") or (xp_user="&user_id&" and xp_score="&score&") or (cp_user1="&user_id&" and cp_score1="&score&") or (cp_user2="&user_id&" and cp_score2="&score&") or (cp_user3="&user_id&" and cp_score3="&score&") or (cp_user4="&user_id&" and cp_score4="&score&") or (cp_user5="&user_id&" and cp_score5="&score&") or (sj_user="&user_id&" and sj_score="&score&") or (hz_user="&user_id&" and hz_score="&score&"))")(0))&"��100��;&nbsp;&nbsp;"

%></td>
  </tr>
</table>
<%end sub
sub ShowSuitType(sidlist)
	if sidlist<>"" and left(sidlist,2)=", " then sidlist=mid(sidlist,2)%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><%set rslosttype=server.createobject("adodb.recordset")
	rslosttype.open "select * from CustomerLostType order by px",conn,1,1
	for lti=1 to rslosttype.recordcount+1
		if lti=rslosttype.recordcount+1 then
			lt_title = "����"
			lt_id = 0
		else
			lt_title = rslosttype("title")
			lt_id = rslosttype("id")
		end if
		if sidlist<>"" then
			lt_money = conn.execute("select sum(s.jixiang_money) from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.customerlosttype="&lt_id&" and s.id in ("&sidlist&")")(0)
			if isnull(lt_money) then lt_money=0
			set rs_ds1 = server.createobject("adodb.recordset")
			set rs_ds3 = server.createobject("adodb.recordset")
			ds1_all = 0
			ds3_all = 0
			rs_ds1.open "select distinct s.id from shejixiadan s inner join kehu k on s.kehu_id=k.id where k.customerlosttype="&lt_id&" and s.id in ("&sidlist&")",conn,1,1
			if not (rs_ds1.eof and rs_ds1.bof) then
				ds1_all = rs_ds1.recordcount
			else
				ds1_all = 0
			end if
			rs_ds1.close
			
			rs_ds3.open "select distinct s.id from (kehu k inner join shejixiadan s on k.id = s.kehu_id) inner join fujia2 f on s.id = f.xiangmu_id where k.customerlosttype="&lt_id&" and "&GetSqlCheckDateString("f.times")&" and s.id in ("&sidlist&")",conn,1,1
			if not (rs_ds3.eof and rs_ds3.bof) then
				ds3_all = rs_ds3.recordcount
			else
				ds3_all = 0
			end if
			rs_ds3.close
		else
			lt_money = 0
			ds1_all = 0
			ds3_all = 0
		end if
		response.write lt_title&"��ϵ��"&lt_money&"Ԫ&nbsp;"
		response.write "δ���ѣ�"&ds1_all-ds3_all&"��&nbsp;&nbsp;&nbsp;"
		if lti mod 4 = 0 then response.write "<br>"
		if not rslosttype.eof then rslosttype.movenext
	next
	rslosttype.close
	set rslosttype = nothing%></td>
  </tr>
</table>
<%end sub

Function ShowWedSignStats(xmlist, uid)
	dim rstype,sqlhs,slhs,strings,sum
	if xmlist<>"" then 
		set rstype=server.createobject("adodb.recordset")
		sqlhs = "select * from hs_signtype order by px asc"
		rstype.open sqlhs,conn,1,1
		do while not rstype.eof
			sum = GetFieldDataBySQL("SELECT sum(vol) FROM hs_signhistory where userid="&uid&" and xiangmu_id in ("&xmlist&") and typeid="&rstype("id")&" and "& GetSqlCheckDateString("idate"),"int",0)
			if isnull(sum) then sum=0
			strings = strings & rstype("title") & sum & "��&nbsp;&nbsp;"
			rstype.movenext
		loop
		rstype.close
		set rstype = nothing
	end if
	ShowWedSignStats = strings
End Function 
%>
<script language="javascript">loadingHidden();</script>
<%
	response.Flush()
end if%>
</body>
</html>

