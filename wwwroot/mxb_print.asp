<!--#include file="connstr.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/SystemWorkflow.Class.asp"-->
<%Response.Buffer=True%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="admin/zxcss.css" rel="stylesheet" type="text/css">
<script language="javascript" src="inc/func.js" type="text/javascript"></script>
<script language="javascript">
function loadingHidden()
{
	eval("document.getElementById(\"loadingimg\").style.display=\"none\"");
}
function loadingShow()
{
	eval("document.getElementById(\"loadingimg\").style.display=\"\"");
}
</script>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
	padding:10px;
}
.style3 {color: #FF0000}
.STYLE4 {font-size: 12pt}
.STYLE5 {font-size: 12}
-->
</style><body topmargin="0" leftmargin="0">
<div id="loadingimg" align="center" style="width:100%; padding-top:100px; float:left; display:none"><img src="Image/loading.gif" width="16" height="16"><br>
  <br>
<div id="loadingtext">������������,���Ե�...</div></div>
<script language="javascript">loadingShow();</script>
<%Response.Flush()%>

<%if request("fromtime")<>"" then
	times=request("fromtime")
else
	times=Date()
end if 

yeard=year(times)
monthd=month(times)
dayd=day(times)

if request("action")="wc" then
	id=request("id")
	if id<>"" and isnumeric(id) then
		conn.execute("update shejixiadan set hs_userid='"&session("userid")&"' where id="&id)	
	end if
end if

function GetCustName(kehu_id)
	dim arr(1)
	If kehu_id="" Or Not IsNumeric(kehu_id) Then
		GetCustName = False
		Exit Function
	End if

	set rskh=conn.execute("select lxpeple,lxpeple2,telephone,telephone2,sex,sex2 from kehu where id="&kehu_id)
	if not rskh.eof then
		if rskh("sex") = "��" then
			arr(0) = rskh("lxpeple")&"<br>"&GetTelNo(rskh("telephone"))
			arr(1) = rskh("lxpeple2")&"<br>"&GetTelNo(rskh("telephone2"))
		else
			arr(0) = rskh("lxpeple2")&"<br>"&GetTelNo(rskh("telephone2"))
			arr(1) = rskh("lxpeple")&"<br>"&GetTelNo(rskh("telephone"))
		end if
		GetCustName = arr
	else
		GetCustName = false
	end if
	rskh.close()
	set rskh = nothing
end function

if session("level")=3 or session("level")=6 or session("level")=7 or session("level")=8 or session("level")=9 or session("level")=10 then
	pageurl = "admin/kehu_mianban.asp"
elseif session("level")=1 then
	pageurl = "kehu_mianban.asp"
elseif session("level")=2 or session("level")=4 or session("level")=5 or session("level")=11 or session("level")=12 or session("level")=13 or session("level")=14 then
	pageurl = "shejishi/kehu_mianban.asp"
end if
%>
<center><%
dim sqlshop,defshopvalue,shopname
if request("chaxun")="" then defshopvalue = GetMultipleShopListValue()
if request("shopid")<>"" then defshopvalue = request("shopid")
if defshopvalue<>"" and not isnull(defshopvalue) then
	sqlshop = " and k.shopid="&defshopvalue
	shopname = GetMultipleShopName(defshopvalue)
else
	sqlshop = ""
	shopname = "ȫ��"
end if

'if not isnull(defshopvalue) and defshopvalue<>"" then
'	if session("shopid")<>0 then
'		sqlshop = " and (k.shopid="&defshopvalue&" or (k.userid in (select username from yuangong where isshare=1)) or (k.userid2 in (select username from yuangong where isshare=1)) or (k.userid3 in (select username from yuangong where isshare=1)))"
'	else
'		sqlshop = " and k.shopid="&defshopvalue
'	end if
'	shopname = GetMultipleShopName(defshopvalue)
'else
'	if session("shopid")<>0 then
'		sqlshop = " and (k.userid in (select username from yuangong where isshare=1)) or (k.userid2 in (select username from yuangong where isshare=1)) or (k.userid3 in (select username from yuangong where isshare=1))"
'	else
'		sqlshop = ""
'	end if
'	shopname = "ȫ��"
'end if

dim CompanyType,IsShowRcsl
CompanyType = GetFieldDataBySQL("select CompanyType from sysconfig","int",0)
IsShowRcsl = GetFieldDataBySQL("select IsShowRcsl from sysconfig","int",0)

response.write shopname
%>&nbsp;<%=yeard%>��<%=monthd%>��<%=dayd%>��
	[
<% select case WEEKDAY(times)
	case 1 
	response.Write "������"
	case 2
	response.Write "����һ"
	case 3
	response.Write "���ڶ�"
	case 4
	response.Write "������"
	case 5
	response.Write "������"
	case 6
	response.Write "������"
	case 7
	response.Write "������"
     end select
	%>
    ]
    <br><br>
<%
response.write ShowPaixiuInfo(times)

dim s,tc_row
s = request.QueryString("s")
select case s
	case "cp"
		show_cp
	case "ky"
		show_ky
	case "xg"
		show_xg
	case "qj"
		show_qj
	case "jhz"
		show_jhz
	case "ls"
		show_ls
	case "pzlf"
		show_pzlf
	case "jhlf"
		show_jhlf
	case "hhz"
		show_hhz
	case else
		if session("level")<>2 and session("level")<>13 then show_cp
		show_ky
		show_xg
		show_qj
		if CompanyType=0 then show_jhz
		show_ls
		show_pzlf
		if CompanyType=0 then show_jhlf
		if CompanyType=0 then show_hhz
end select
sub show_cp()%>
    <span class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=cp&shopid=<%=defshopvalue%>" title="��ҳ��ʾ��ر�">��ر�</a></span><br>
  <br>
</center>
<table width="100%" height="25" border="1" cellspacing="0">
  <tr  bgcolor="#66FFFF" align="center" height="25">
    <td width="30" height="25">���</td>
    <td>�ֶ�����<br>
      ����/ʱ��<br>
����</td>
    <td><%=GetAppellation(3, false)%></td>
    <td><%=GetAppellation(4, true)%></td>
    <td>��ϵ����<br>
      ��ϵ����</td>
    <td>��δ��<br>
      ��ϵ��</td>
    <td>����<br>
    <font color=red>����</font></td>
    <td>Ԥ��Ӱ<br>
    <font color=red>ʵ��Ӱ</font></td>
    <td>Ԥ<%=GetWorkName("hz")%>1<br>
        <font color=red>ʵ<%=GetWorkName("hz")%>1</font></td>
    <td>Ԥ<%=GetWorkName("hz")%>2<br>
        <font color=red>ʵ<%=GetWorkName("hz")%>2</font></td>
    <td>Ԥ����<br>
      <font color=red>ԤѡƬ</font></td>
    <td>&nbsp;����</td>
    <td>���<br>
      <font color=red>��������</font></td>
    <td>���<%if IsShowRcsl=1 then response.write "<br><font color=red>����</font>"%></td>
    <td>Ԥ���<br>
    <font color=red>ʵ���</font></td>
    <td>ѡƬʱ��</td>
    <td>&nbsp;��ע</td>
  </tr>
<%
dim c,sqlpz
c=0
sqlpz="select * from (SELECT s.ID, s.danhao, s.jixiang, s.kehu_id, s.jixiang_money, pz_time, s.pz, s.hz_time, s.userid, s.yx_cp_name, s.yx_cp_name2, s.yx_cp_name3, s.cp_name, s.cp_name2, s.cp_name2, s.cp_name3, s.cp_name4, s.cp_name5, s.cpVolume, s.yx_hz_name, s.yx_hz_name2, s.hz_name2nd, s.hz_name, s.yx_hzzl_name, s.hz_name2, s.yx_ky_name, s.yunyong, s.sl, s.yx_cpzl_name, s.yx_jhlf_name, s.hs_userid, s.kj_time, s.yx_cp_memo,s.yx_cp_memo2,s.yx_cp_memo3,s.tc_pz_time,s.tc_pz_time2,s.tc_pz_time3,sl2,rcsl,1 as sig FROM shejixiadan s LEFT JOIN kehu k ON s.kehu_id = k.ID where s.pz_time=#"&times&"#"&sqlshop&" union all SELECT s.ID, s.danhao, s.jixiang, s.kehu_id, s.jixiang_money, pz_time2 as pz_time, pz2 as pz, s.hz_time, s.userid, s.yx_cp_name, s.yx_cp_name2, s.yx_cp_name3, s.cp_name, s.cp_name2, s.cp_name2, s.cp_name3, s.cp_name4, s.cp_name5, s.cpVolume, s.yx_hz_name, s.yx_hz_name2, s.hz_name2nd, s.hz_name, s.yx_hzzl_name, s.hz_name2, s.yx_ky_name, s.yunyong, s.sl, s.yx_cpzl_name, s.yx_jhlf_name, s.hs_userid, s.kj_time, s.yx_cp_memo,s.yx_cp_memo2,s.yx_cp_memo3,s.tc_pz_time,s.tc_pz_time2,s.tc_pz_time3,sl2,rcsl,2 as sig FROM shejixiadan s LEFT JOIN kehu k ON s.kehu_id = k.ID where pz_time2=#"&times&"#"&sqlshop&" union all SELECT s.ID, s.danhao, s.jixiang, s.kehu_id, s.jixiang_money, pz_time3 as pz_time, pz3 as pz, s.hz_time, s.userid, s.yx_cp_name, s.yx_cp_name2, s.yx_cp_name3, s.cp_name, s.cp_name2, s.cp_name2, s.cp_name3, s.cp_name4, s.cp_name5, s.cpVolume, s.yx_hz_name, s.yx_hz_name2, s.hz_name2nd, s.hz_name, s.yx_hzzl_name, s.hz_name2, s.yx_ky_name, s.yunyong, s.sl, s.yx_cpzl_name, s.yx_jhlf_name, s.hs_userid, s.kj_time, s.yx_cp_memo,s.yx_cp_memo2,s.yx_cp_memo3,s.tc_pz_time,s.tc_pz_time2,s.tc_pz_time3,sl2,rcsl,3 as sig FROM shejixiadan s LEFT JOIN kehu k ON s.kehu_id = k.ID where pz_time3=#"&times&"#"&sqlshop&") order by pz"
set rs=server.CreateObject("adodb.recordset")
'"select s.* from shejixiadan s inner join kehu k on s.kehu_id=k.id where (s.pz_time=#"&times&"# or s.pz_time2=#"&times&"#)"&sqlshop&" order by s.pz_time desc, pz asc"
rs.open sqlpz,conn,1,1

while not rs.eof
	c=c+1
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	if isnull(money1) then money1=0
  	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id"))(0)
	if isnull(money2) then money2=0
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id"))(0)
	if isnull(money3) then money3=0
	money4=conn.execute("select jixiang_money from shejixiadan where id="&rs("id"))(0)
	if isnull(money4) then money4=0
	money5=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id"))(0)
	if isnull(money5) then money5=0
	if rs("sig")=1 then
		tc_row = GetDatecolor(rs("tc_pz_time"), "")
		yx_cp_memo = rs("yx_cp_memo")
	elseif rs("sig")=2 then
		tc_row = GetDatecolor(rs("tc_pz_time2"), "")
		yx_cp_memo = rs("yx_cp_memo2")
	elseif rs("sig")=3 then
		tc_row = GetDatecolor(rs("tc_pz_time3"), "")
		yx_cp_memo = rs("yx_cp_memo3")
	end if
	if tc_row="" or isnull(tc_row) then tc_row="#ffffff"
%>  
  <tr  bgcolor="<%=tc_row%>" onClick="javascript:openEditScript('<%=pageurl%>?id=<%=rs("id")%>',450,500)" style="cursor:hand">
  <!--onMouseOver="this.bgColor='#FF9966'" onMouseOut="this.bgColor='#FFFFFF'" -->
    <td  height="25" align="center"><%=c%></td>
    <td  height="25" align="center"><%
	if rs("danhao")<>"" and not isnull(rs("danhao")) then
		response.write rs("danhao")&"<br>"
	end if
	response.write rs("id") & "/" & rs("pz")
	  response.write "<br>" & rs("hz_time")%></td>
    <td align="center">
    <%
	arr = GetCustName(rs("kehu_id"))
	if not isarray(arr) then
		redim arr(1)
		arr(0) = "N/A"
		arr(1) = "N/A"
	end if
	response.write arr(0)
	%></td>
    <td align="center"><%=arr(1)%></td>
    <td align="center"><%dim rsjx
	set rsjx = conn.execute("select companytype.companytype,jixiang.jixiang from companytype inner join jixiang on companytype.id=jixiang.type where jixiang.id="&rs("jixiang"))
	if not (rsjx.eof and rsjx.bof) then
		response.write rsjx("companytype")&"<br />"&rsjx("jixiang")
	else
		response.write "&nbsp;"
	end if
	rsjx.close
	set rsjx=nothing%></td>
    <td align="center"><%=money1+money2+money3+money4-money5%><br>
    <%=rs("jixiang_money")%></td>
    <td align="center"><%response.write GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid")&"'","str","")
	response.write "<br>"
	response.write "<font color=red>"
	if rs("yx_cpzl_name")<>"" and not isnull(rs("yx_cpzl_name")) then
		response.write rs("yx_cpzl_name")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"%></td>
    <td align="center"><%
	if rs("yx_cp_name")<>"" and not isnull(rs("yx_cp_name")) then response.write rs("yx_cp_name")
	if rs("yx_cp_name2")<>"" and not isnull(rs("yx_cp_name2")) then response.write "/"&rs("yx_cp_name2")
	if rs("yx_cp_name3")<>"" and not isnull(rs("yx_cp_name3")) then response.write "/"&rs("yx_cp_name3")
	
	response.write "<font color=red>"
	if rs("cp_name")<>"" and not isnull(rs("cp_name")) then response.write "<br>"&rs("cp_name")
	if rs("cp_name2")<>"" and not isnull(rs("cp_name2")) then response.write "/"&rs("cp_name2")
	if rs("cp_name3")<>"" and not isnull(rs("cp_name3")) then response.write "/"&rs("cp_name3")
	if rs("cp_name4")<>"" and not isnull(rs("cp_name4")) then response.write "/"&rs("cp_name4")
	if rs("cp_name5")<>"" and not isnull(rs("cp_name5")) then response.write "/"&rs("cp_name5")
	response.write "</font>"
	%>&nbsp;</td>
    <td align="center"><%if rs("yx_hz_name")<>"" and not isnull(rs("yx_hz_name")) then
		response.write rs("yx_hz_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	response.write "<font color=red>"
	if rs("hz_name")<>"" and not isnull(rs("hz_name")) then
		response.write rs("hz_name")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"
	%></td>
    <td align="center"><%if rs("yx_hz_name2")<>"" and not isnull(rs("yx_hz_name2")) then
		response.write rs("yx_hz_name2")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	response.write "<font color=red>"
	if rs("hz_name2nd")<>"" and not isnull(rs("hz_name2nd")) then
		response.write rs("hz_name2nd")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"
	%></td>
    <td align="center"><%if rs("yx_hzzl_name")<>"" and not isnull(rs("yx_hzzl_name")) then
		response.write rs("yx_hzzl_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	response.write "<font color=red>"
	if rs("yx_ky_name")<>"" and not isnull(rs("yx_ky_name")) then
		response.write rs("yx_ky_name")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"
	%></td>
    <td><%
	yunyong_list = rs("yunyong")
'	arr_yunyong = split(yunyong_list,", ")
'	for k = 0 to ubound(arr_yunyong)
'		wj_flag = conn.execute("select iswj from yunyong_type where id in (select type_id from yunyong where id="&arr_yunyong(k)&")")(0)
'		if wj_flag = 1 then
'			response.write "&nbsp;"&conn.execute("select yunyong from yunyong where id="&arr_yunyong(k))(0)
'			if k<ubound(arr_yunyong) then response.write "<br>"
'		end if
'	next
	If yunyong_list="" Or IsNull(yunyong_list) Then
		response.write "��������"
	Else
		set rswj = conn.execute("select * from yunyong where id in ("&yunyong_list&") and type_id in (select id from yunyong_type where iswj=1) order by id")
		do while not rswj.eof
			response.write "&nbsp;"&rswj("yunyong")
			rswj.movenext
			if not rswj.eof then response.write "<br>"
		loop
		rswj.close
		set rswj = Nothing
	End If
	%>&nbsp;</td>
    <td align="center"><%if rs("yunyong")<>"" and not isnull(rs("yunyong")) then
		arryy = split(rs("yunyong"),", ")
		arrsl = split(rs("sl"),", ")
		lfcounts = 0
		for i = 0 to ubound(arrsl)
			set rslf = conn.execute("select * from yunyong where id="&arryy(i))
			if not rslf.eof then
				if rslf("type3") = 1 then
					if lfcounts>0 then response.write "<br>" 
					response.write rslf("yunyong")&"["&arrsl(i)&"]"
					lfcounts=lfcounts+1
				end if
			end if
			rslf.close
			set rslf=nothing
		next
	end if
	response.write "<br><font color='red'>"&rs("cpVolume")&"��</font>"%></td>
    <td align="center"><%response.write rs("sl2")
	if IsShowRcsl=1 then response.write "<br><font color=red>"&rs("rcsl")&"<br>"%></td>
    <td align="center"><%if rs("yx_jhlf_name")<>"" and not isnull(rs("yx_jhlf_name")) then
		response.write rs("yx_jhlf_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	response.write "<font color=red>"
	if rs("hs_userid")<>"" and not isnull(rs("hs_userid")) then
		response.write GetFieldDataBySQL("select peplename from yuangong where username='"&rs("hs_userid")&"'","str","")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"
	%></td>
    <td align="center"><%if not isnull(rs("kj_time")) then
		response.write rs("kj_time")
	else
		response.write "&nbsp;"
	end if
	%></td>
    <td><%=yx_cp_memo%>&nbsp;</td>
  </tr>
<%
rs.movenext
wend
rs.close
set rs=nothing
%> 
</table>
<%end sub
sub show_ky()
	c=0%>
<p align="center" class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=ky&shopid=<%=defshopvalue%>" title="��ҳ��ʾѡƬ��">ѡƬ��</a></p>
<table width="100%" border="1" cellspacing="0">
  <tr bgcolor="#66FFFF" align="center" height="25">
    <td width="50">���</td>
    <td width="12%" height="25">�ֶ�����<br>
      ����/ʱ��<br>
����</td>
    <td>��ϵ<br>
    ��ϵ��</td>
    <td><%=GetAppellation(3, false)%></td>
    <td><%=GetAppellation(4, true)%></td>
    <td>��δ��<br>
      ѡƬ��</td>
    <td bgcolor="#66FFFF">����</td>
    <td bgcolor="#66FFFF">��Ӱʦ/<%=GetDutyName(5)%><br>
      ȡ��ʱ��</td>
    <td bgcolor="#66FFFF">���ʦ</td>
    <td>�� ɫ</td>
    <td>ԤѡƬ<br>
        <font color=red>ʵѡƬ</font></td>
    <td>��ע</td>
  </tr>
  <%
 year1=year(times)
 day1=day(times)
 month1=month(times)
set rs=server.CreateObject("adodb.recordset")
'rs.open "select * from shejixiadan where lc_ky=#"&times&"#  order by hz_time desc",conn,1,1
'rs.open "select * from shejixiadan where day(lc_ky)="&day1&" and month(lc_ky)="&month1&" and year(lc_ky)="&year1&" order by hz_time desc",conn,1,1
rs.open "select s.* from shejixiadan s inner join kehu k on s.kehu_id=k.id where s.kj_time=#"&times&"#"&sqlshop&" order by kj asc",conn,1,1
while not rs.eof
	c=c+1
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	if isnull(money1) then money1=0
  	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id"))(0)
	if isnull(money2) then money2=0
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id"))(0)
	if isnull(money3) then money3=0
	money4=conn.execute("select jixiang_money from shejixiadan where id="&rs("id"))(0)
	if isnull(money4) then money4=0
	money5=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id"))(0)
	if isnull(money5) then money5=0
	
	tc_row = GetDatecolor(rs("tc_kj_time"), "")
	if tc_row="" or isnull(tc_row) then tc_row="#ffffff"
%>
  <tr align="center" bgcolor="<%=tc_row%>" onClick="javascript:openEditScript('<%=pageurl%>?id=<%=rs("id")%>',450,500)" style="cursor:hand">
    <td align="center"><%=c%></td>
    <td  height="25" align="center">&nbsp;
      <%
	if rs("danhao")<>"" and not isnull(rs("danhao")) then
		response.write rs("danhao")&"<br>"
	end if
	%>
      <%=rs("id")%>/<%=rs("kj")%><br>
    <%=rs("hz_time")%></td>
    <td align="center"><%=GetFieldDataBySQL("select jixiang from jixiang where id="&rs("jixiang"),"str","&nbsp;")%>
      <br>
    <%=rs("jixiang_money")%></td>
    <td align="center"><%
	arr = GetCustName(rs("kehu_id"))
	if not isarray(arr) then
		redim arr(1)
		arr(0) = "N/A"
		arr(1) = "N/A"
	end if
	response.write arr(0)
	%></td>
    <td align="center">&nbsp;<%=arr(1)%></td>
    <td align="center"><%response.write money1+money2+money3+money4-money5
	hqmoney = conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	if isnull(hqmoney) then hqmoney=0
	response.write "<br>"&hqmoney
	%></td>
    <td align="center"><%=getfielddatabysql("select peplename from yuangong where username='"&rs("userid")&"'","str","&nbsp;")%></td>
    <td align="center"><%=rs("cp_name")%>/<%=rs("hz_name")%><br><%=rs("qj_time")%></td>
    <td align="center"><%
	if rs("sj_name")<>"" and not isnull(rs("sj_name")) then
		if isnull(rs("lc_sj")) then
			response.write "<font color=red>"&rs("sj_name")&"<br>(δ���)</font>"
		else
			response.write rs("sj_name")
		end if
	else
		response.write "&nbsp;"
	end if
	%></td>
    <td align="center"><%=RS("xp_name")%>&nbsp;</td>
    <td align="center"><%if rs("yx_ky_name")<>"" and not isnull(rs("yx_ky_name")) then
		response.write rs("yx_ky_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	response.write "<font color=red>"
	if rs("ky_name")<>"" and not isnull(rs("ky_name")) then
		response.write rs("ky_name")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"
	%></td>
    <td align="left"><%=rs("yx_cp_memo")%>&nbsp;</td>
  </tr>
  <%
rs.movenext
wend
rs.close
set rs=nothing
%>
</table>
<%end sub
sub show_xg()%>
<p align="center" class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=xg&shopid=<%=defshopvalue%>" title="��ҳ��ʾ�����">�����</a></p>
<table width="100%" border="1" cellspacing="0">
  <tr bgcolor="#66FFFF" align="center" height="25">
    <td height="25">�ֶ�����<br>����/ʱ��<br>
����</td>
    <td><%=GetAppellation(3, false)%></td>
    <td><%=GetAppellation(4, true)%></td>
    <td>&nbsp;��δ��</td>
    <td>ȡ��ʱ��</td>
    <td>���ʦ</td>
    <td>���淽ʽ</td>
    <td>Ԥ����<br>
        <font color=red>ʵ����</font></td>
  </tr>
  <%
 year1=year(times)
 day1=day(times)
 month1=month(times)
set rs=server.CreateObject("adodb.recordset")
'rs.open "select * from shejixiadan where lc_ky=#"&times&"#  order by hz_time desc",conn,1,1
'rs.open "select * from shejixiadan where day(lc_ky)="&day1&" and month(lc_ky)="&month1&" and year(lc_ky)="&year1&" order by hz_time desc",conn,1,1
rs.open "select s.* from shejixiadan s inner join kehu k on s.kehu_id=k.id where s.xg_time=#"&times&"#"&sqlshop&" order by s.xg_time desc, xg asc",conn,1,1
while not rs.eof
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	if isnull(money1) then money1=0
  	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id"))(0)
	if isnull(money2) then money2=0
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id"))(0)
	if isnull(money3) then money3=0
	money4=conn.execute("select jixiang_money from shejixiadan where id="&rs("id"))(0)
	if isnull(money4) then money4=0
	money5=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id"))(0)
	if isnull(money5) then money5=0
	
	tc_row = GetDatecolor(rs("tc_xg_time"), "")
	if tc_row="" or isnull(tc_row) then tc_row="#ffffff"
%>
  <tr align="center" bgcolor="<%=tc_row%>" onClick="javascript:openEditScript('<%=pageurl%>?id=<%=rs("id")%>',450,500)" style="cursor:hand">
    <td  height="25"><%
	if rs("danhao")<>"" and not isnull(rs("danhao")) then
		response.write rs("danhao")&"<br>"
	end if
	%>
      <%=rs("id")%>/<%=rs("xg")%><br>
    <%=rs("hz_time")%></td>
    <td><%
	arr = GetCustName(rs("kehu_id"))
	if not isarray(arr) then
		redim arr(1)
		arr(0) = "N/A"
		arr(1) = "N/A"
	end if
	response.write arr(0)
	%></td>
    <td><%=arr(1)%></td>
    <td>&nbsp;<%=money1+money2+money3+money4-money5%></td>
    <td>&nbsp;<%=rs("qj_time")%></td>
    <td>&nbsp;<%if rs("sj_name")<>"" and not isnull(rs("sj_name")) then
		if isnull(rs("lc_sj")) then
			response.write "<font color=red>"&rs("sj_name")&"<br>(δ���)</font>"
		else
			response.write rs("sj_name")
		end if
	else
		response.write "&nbsp;"
	end if%></td>
    <td><%if rs("xg_opt")=0 then
		response.write "�ڲ�����"
	else
		response.write "���˿���"
	end if%></td>
    <td><%if rs("yx_xg_name")<>"" and not isnull(rs("yx_xg_name")) then
		response.write rs("yx_xg_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	response.write "<font color=red>"
	if rs("xg_name")<>"" and not isnull(rs("xg_name")) then
		response.write rs("xg_name")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"
	%></td>
  </tr>
  <%
rs.movenext
wend
rs.close
set rs=nothing
%>
</table>
<%end sub
sub show_qj()%>
<p align="center" class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=qj&shopid=<%=defshopvalue%>" title="��ҳ��ʾȡ����">ȡ����</a></p>
<table width="100%" border="1" cellspacing="0">
  <tr bgcolor="#66FFFF" align="center" height="25">
    <td height="25">�ֶ�����<br>����/ʱ��<br>
����</td>
    <td><%=GetAppellation(3, false)%></td>
    <td><%=GetAppellation(4, true)%></td>
    <td>&nbsp;��δ��</td>
    <td>�ӵ�ʱ��</td>
    <td>&nbsp;����</td>
    <td>��ǰ����</td>
    <td>&nbsp;��ע</td>
  </tr>
  <%
dim clsWorkflow
set clsWorkflow = new SystemWorkflow
clsWorkflow.DBConnection = conn
clsWorkflow.LoadInstance(false)  
  
set rs=server.CreateObject("adodb.recordset")
sqlqj="select * from (SELECT s.ID, s.danhao, s.kehu_id, s.jixiang_money, s.qj_time, s.qj, s.userid, 1 as sig, s.lc_wc, s.hz_time, s.tc_qj_time2 as tc_qj_time, s.times FROM shejixiadan s LEFT JOIN kehu k ON s.kehu_id = k.ID where s.qj_time=#"&times&"#"&sqlshop&" union all SELECT s.ID, s.danhao, s.kehu_id, s.jixiang_money, s.qj_time2 as qj_time, qj2 as qj, s.userid, 2 as sig, s.lc_wc, s.hz_time, s.tc_qj_time, s.times FROM shejixiadan s LEFT JOIN kehu k ON s.kehu_id = k.ID where qj_time2=#"&times&"#"&sqlshop&") order by hz_time desc"
rs.open sqlqj,conn,1,1
while not rs.eof
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	if isnull(money1) then money1=0
  	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id"))(0)
	if isnull(money2) then money2=0
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id"))(0)
	if isnull(money3) then money3=0
	money4=conn.execute("select jixiang_money from shejixiadan where id="&rs("id"))(0)
	if isnull(money4) then money4=0
	money5=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id"))(0)
	if isnull(money5) then money5=0
	
	tc_row = GetDatecolor(rs("tc_qj_time"), "")
	if tc_row="" or isnull(tc_row) then tc_row="#ffffff"
%>
  <tr align="center" bgcolor="<%=tc_row%>" onClick="javascript:openEditScript('<%=pageurl%>?id=<%=rs("id")%>',450,500)" style="cursor:hand">
    <td  height="25">&nbsp;
      <%
	if rs("danhao")<>"" and not isnull(rs("danhao")) then
		response.write rs("danhao")&"<br>"
	end if
	%>
      <%=rs("id")%>/<%=rs("qj")%><br>
    <%=rs("hz_time")%></td>
    <td><%
	arr = GetCustName(rs("kehu_id"))
	if not isarray(arr) then
		redim arr(1)
		arr(0) = "N/A"
		arr(1) = "N/A"
	end if
	response.write arr(0)
	%></td>
    <td><%=arr(1)%></td>
    <td>&nbsp;<%=money1+money2+money3+money4-money5%></td>
    <td><%=datevalue(rs("times"))%></td>
    <td>&nbsp;<%
	response.write GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid")&"'","str","")%></td>
    <td><%if not isnull(rs("lc_wc")) then
		response.write "ȫ�����"
	else
		dim currentwork
		currentwork = clsWorkflow.GetCurrentWork(rs("id"), 2)
		if isnull(currentwork) or currentwork="" then
			response.write "δ��ʼ"
		else
			response.write currentwork & "���"
		end if
	end if
	
	%>    </td>
    <td>&nbsp;</td>
  </tr>
  <%
rs.movenext
wend
rs.close
set rs=nothing
set clsWorkflow=nothing
%>
</table>
<%end sub
sub show_jhz()%>
<p align="center" class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=jhz&shopid=<%=defshopvalue%>" title="��ҳ��ʾ��黯ױ��">��黯ױ��</a> </p>
<table width="100%" border="1" cellspacing="0">
  <tr  bgcolor="#66FFFF" align="center" height="25">
    <td height="25">�ֶ�����<br>����/ʱ��<br>
����</td>
    <td><%=GetAppellation(3, false)%></td>
    <td><%=GetAppellation(4, true)%></td>
    <td>��δ��</td>
    <td>&nbsp;����</td>
    <td>Ԥ���ױ<br>
        <font color=red>ʵ���ױ</font></td>
    <td>Ԥ���ʦ<br>
      <font color=red>ʵ���ʦ</font></td>
    <td>����б�</td>
    <td>˾��</td>
    <td>����</td>
    <td>�鳵</td>
    <td>�ֳ�ʱ��</td>
    <td>��ע</td>
  </tr>
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select s.* from shejixiadan s inner join kehu k on s.kehu_id=k.id where s.hz_time=#"&times&"#"&sqlshop&" order by s.hz",conn,1,1
while not rs.eof
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	if isnull(money1) then money1=0
  	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id"))(0)
	if isnull(money2) then money2=0
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id"))(0)
	if isnull(money3) then money3=0
	money4=conn.execute("select jixiang_money from shejixiadan where id="&rs("id"))(0)
	if isnull(money4) then money4=0
	money5=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id"))(0)
	if isnull(money5) then money5=0
	
	tc_row = GetDatecolor(rs("tc_hz_time"), "")
	if tc_row="" or isnull(tc_row) then tc_row="#ffffff"
%>  
  <tr align="center" bgcolor="<%=tc_row%>" onClick="javascript:openEditScript('<%=pageurl%>?id=<%=rs("id")%>',450,500)" style="cursor:hand">
    <td  height="25"><%
	if rs("danhao")<>"" and not isnull(rs("danhao")) then
		response.write rs("danhao")&"<br>"
	end if
	%>
      <%=rs("id")%>/<%=rs("hz")%></td>
    <td><%
	arr = GetCustName(rs("kehu_id"))
	if not isarray(arr) then
		redim arr(1)
		arr(0) = "N/A"
		arr(1) = "N/A"
	end if
	response.write arr(0)
	%></td>
    <td><%=arr(1)%></td>
    <td>&nbsp;<%=money1+money2+money3+money4-money5%></td>
    <td>&nbsp;<%
	response.write GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid")&"'","str","")%></td>
    <td><%if rs("yx_jhz_name")<>"" and not isnull(rs("yx_jhz_name")) then
		response.write rs("yx_jhz_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	response.write "<font color=red>"
	if rs("hz_userid")<>"" and not isnull(rs("hz_userid")) then
		response.write GetFieldDataBySQL("select peplename from yuangong where username='"&rs("hz_userid")&"'","str","")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"
	%></td>
    <td><%if rs("yx_jhlf_name")<>"" and not isnull(rs("yx_jhlf_name")) then
		response.write rs("yx_jhlf_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	response.write "<font color=red>"
	if rs("hs_userid")<>"" and not isnull(rs("hs_userid")) then
		response.write GetFieldDataBySQL("select peplename from yuangong where username='"&rs("hs_userid")&"'","str","")
	else
		response.write "&nbsp;"
	end if
	response.write "</font>"
	%></td>
    <td><table width="95%" border="0" cellspacing="0" cellpadding="0">
      <%
		set rshs = server.CreateObject("adodb.recordset")
		sqlhs = "select d.* from chuzhu_details d right join chuzhu_jilu j on d.orderid=j.id where j.xiangmu_id="&rs("id")
		rshs.open sqlhs,conn,1,1
		do while not rshs.eof
		  	set rshsd = conn.execute("select num from huensha where id="&rshs("AnnexWedID"))
		  	if not (rshsd.eof and rshsd.bof) then
			  %>
			<tr>
			  <td><%=rshsd("num")%></td>
			  <td>[<%=rshs("Volume")%>]</td>
			</tr>
			<%
			end if
			rshsd.close
			set rshsd = nothing
			rshs.movenext
		loop
		rshs.close
		set rshs = nothing
		%>
      </table></td>
    <td><%=rs("emcee_name")%>&nbsp;</td>
    <td><%=rs("rec_name")%>&nbsp;</td>
    <td><%=rs("car_info")%>&nbsp;</td>
    <td><%=rs("locale_time")%>&nbsp;</td>
    <td><%=rs("yx_cp_memo")%>&nbsp;</td>
  </tr>
<%
rs.movenext
wend
rs.close
set rs=nothing
%>   
</table>
<%end sub
sub show_ls()%>
<p align="center" class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=ls&shopid=<%=defshopvalue%>" title="��ҳ��ʾ��ɢ���ѱ�">��ɢ���ѱ�</a></p>
<table width="100%" border="1" cellspacing="0">
  <tr  bgcolor="#66FFFF" align="center" height="25">
    <td height="25">����</td>
    <td>�ͻ�</td>
    <td>�绰</td>
    <td>������1</td>
    <td>������2</td>
    <td>��Ŀ</td>
    <td>���</td>
    <td>&nbsp;��ע</td>
  </tr>
  <%
set rstx=server.CreateObject("adodb.recordset")
rstx.open "select distinct bm_id from goumai_jilu where not isnull(txtimes) and txtimes=#"&times&"# order by bm_id",conn,1,1
do while not rstx.eof
	set rs=conn.execute("select * from goumai_jilu where bm_id='"&rstx("bm_id")&"' order by id")
	bm_id=rs("bm_id")
	kehu_name=rs("kehu_name")
	telephone=rs("telephone")
	beizhu=rs("beizhu")
	counts=rs("counts")
%>
  <tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#FF9966'" onMouseOut="this.bgColor='#FFFFFF'" onClick="javascript:openEditScript('admin/print_frame.asp?url=save_print.asp?counts=<%=counts%>',750,350)" style="cursor:hand">
    <td  height="25" align="center"><%="Y"&bm_id%></td>
    <td align="center">&nbsp;<%=kehu_name%></td>
    <td align="center">&nbsp;<%=telephone%></td>
    <td align="center"><%=GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid")&"'","str","&nbsp;")%></td>
    <td align="center"><%=GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid2")&"'","str","&nbsp;")%></td>
    <td>&nbsp;<%do while not rs.eof
		response.write GetFieldDataBySQL("select xiangmu from save_type where id="&rs("xiangmu_id"),"str","")&"["&rs("sl")&"]."
		rs.movenext
	loop
	%></td>
    <td align="center"><%dim qk
	qk=GetScattereArrearage(counts)
	if qk>0 then 
		response.write formatnumber(qk,1,0,0,0)
	else
		response.write "0.0"
	end if%></td>
    <td>&nbsp;<%=beizhu%></td>
  </tr>
  <%
  	
	rs.close()
rstx.movenext
loop
rstx.close
set rstx=nothing
%>
</table>
<%end sub
sub show_pzlf()%>
<p align="center" class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=pzlf&shopid=<%=defshopvalue%>" title="��ҳ��ʾ���������">���������</a> </p>
<table width="100%" border="1" cellspacing="0">
  <tr  bgcolor="#66FFFF" align="center" height="25">
    <td width="13%" height="25">�ֶ�����<br>
    ����/ʱ��<br>
����</td>
    <td width="13%"><%=GetAppellation(3, false)%></td>
    <td width="13%"><%=GetAppellation(4, true)%></td>
    <td width="15%">Ԥ���ʦ<br>
    <font color="#FF0000">�������</font></td>
    <td width="15%">ʱ���</td>
    <td width="15%"><%=GetWorkName("hz")%></td>
    <td width="35%">&nbsp;��ע</td>
  </tr>
  <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where pzlf_time=#"&times&"# order by pzlf",conn,1,1
do while not rs.eof
	tc_row = GetDatecolor(rs("tc_pzlf_time"), "")
	if tc_row="" or isnull(tc_row) then tc_row="#ffffff"
%>
  <tr bgcolor="<%=tc_row%>" onClick="javascript:openEditScript('<%=pageurl%>?id=<%=rs("id")%>',450,500)" style="cursor:hand">
    <td  height="25" align="center"><%
	if rs("danhao")<>"" and not isnull(rs("danhao")) then
		response.write rs("danhao")&"<br>"
	end if
	response.write rs("id")&"/"&rs("pzlf")%><br>
    <%=rs("hz_time")%></td>
    <td align="center"><%
	arr = GetCustName(rs("kehu_id"))
	if not isarray(arr) then
		redim arr(1)
		arr(0) = "N/A"
		arr(1) = "N/A"
	end if
	response.write arr(0)
	%></td>
    <td align="center"><%=arr(1)%></td>
    <td align="center"><%
	if rs("yx_hzzl_name")<>"" and not isnull(rs("yx_hzzl_name")) then
		response.write rs("yx_hzzl_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	
	response.write "<font color='red'>"
	set rs_pzlf = conn.execute("select * from hs_pzyd where xiangmu_id="&rs("id"))
	if rs_pzlf.eof and rs_pzlf.bof then
		response.write rs("pzlf_time")
		pzlf = trim(rs("pzlf"))
		if pzlf<>"" and isnumeric(pzlf) then
			if pzlf>0 and pzlf<12 then
				sjd = 0
			else
				sjd = 1
			end if
		else
			sjd = -1
		end if
	else
		set rspzlf = conn.execute("select peplename from yuangong where id="&rs_pzlf("adminid"))
		if not rspzlf.eof then
			response.write rspzlf(0)
		else
			response.write "&nbsp;"
		end if
		rspzlf.close
		set rspzlf = nothing
		sjd = rs_pzlf("cpamorpm")
	end if 
	rs_pzlf.close
	set rs_pzlf = nothing
	response.write "</font>"
	%></td>
    <td align="center"><%
	if sjd=0 then
		response.write "����"
	elseif sjd=1 then
		response.write "����"
	end if
	if pzlf<>"" and isnumeric(pzlf) then
		response.write pzlf&"��"
	end if
	%></td>
    <td align="center">&nbsp;<%=rs("hz_name")%></td>
    <td>&nbsp;</td>
  </tr>
  <%
	rs.movenext
loop
rs.close
set rs=nothing
%>
</table>
<%end sub
sub show_jhlf()%>
<p align="center" class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=jhlf&shopid=<%=defshopvalue%>" title="��ҳ��ʾ��������">��������</a> </p>
<table width="100%" border="1" cellspacing="0">
  <tr  bgcolor="#66FFFF" align="center" height="25">
    <td height="25">�ֶ�����<br>ϵͳ����<br>
����</td>
    <td><%=GetAppellation(3, false)%></td>
    <td><%=GetAppellation(4, true)%></td>
    <td>Ԥ���ʦ<br>
    <font color="#FF0000">�������</font></td>
    <td>ʱ���</td>
    <td>����ʱ��</td>
    <td>����ʱ��</td>
    <td>��黯ױ</td>
    <td width="200">&nbsp;��ע</td>
  </tr>
  <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where jhlf_time=#"&times&"# order by id",conn,1,1
do while not rs.eof
	tc_row = GetDatecolor(rs("tc_jhlf_time"), "")
	if tc_row="" or isnull(tc_row) then tc_row="#ffffff"
	jhlf = trim(rs("jhlf"))
%>
  <tr bgcolor="<%=tc_row%>" onClick="javascript:openEditScript('<%=pageurl%>?id=<%=rs("id")%>',450,500)" style="cursor:hand">
    <td  height="25" align="center"><%
	if rs("danhao")<>"" and not isnull(rs("danhao")) then
		response.write rs("danhao")&"<br>"
	end if
	%>
      <%=rs("id")%><br>
    <%=rs("hz_time")%></td>
    <td align="center"><%
	arr = GetCustName(rs("kehu_id"))
	if not isarray(arr) then
		redim arr(1)
		arr(0) = "N/A"
		arr(1) = "N/A"
	end if
	response.write arr(0)
	%></td>
    <td align="center"><%=arr(1)%></td>
    <td align="center"><%
	if rs("yx_jhlf_name")<>"" and not isnull(rs("yx_jhlf_name")) then
		response.write rs("yx_jhlf_name")
	else
		response.write "&nbsp;"
	end if
	response.write "<br>"
	
	response.write "<font color='red'>"
	set rs_jhlf = conn.execute("select * from chuzhu_jilu where xiangmu_id="&rs("id"))
	if rs_jhlf.eof and rs_jhlf.bof then
		response.write rs("jhlf_time")
		if jhlf<>"" and isnumeric(jhlf) then
			if jhlf>0 and jhlf<12 then
				sjd = 0
			else
				sjd = 1
			end if
		else
			sjd = -1
		end if
	else
		response.write GetFieldDataBySQL("select peplename from yuangong where id="&rs_jhlf("userid"),"str","")
		sjd = rs_jhlf("cpamorpm")
		indate = rs_jhlf("indate")&"&nbsp;"
		outdate = rs_jhlf("outdate")&"&nbsp;"
	end if 
	rs_jhlf.close
	set rs_jhlf = nothing
	response.write "</font>"
	%></td>
    <td align="center"><%
	if sjd=0 then
		response.write "����"
	elseif sjd=1 then
		response.write "����"
	else
		response.write "&nbsp;"
	end if
	if jhlf<>"" and isnumeric(jhlf) then
		response.write jhlf&"��"
	end if
	%></td>
    <td align="center"><%=indate%>&nbsp;</td>
    <td align="center"><%=outdate%>&nbsp;</td>
    <td align="center">&nbsp;<%if rs("hz_userid")<>"" and not isnull(rs("hz_userid")) then
		response.write GetFieldDataBySQL("select peplename from yuangong where username='"&rs("hz_userid")&"'","str","")
	end if
	%></td>
    <td>&nbsp;</td>
  </tr>
  <%
	rs.movenext
loop
rs.close
set rs=nothing
%>
</table>
<%end sub
sub show_hhz()%>
<p align="center" class="STYLE4"><a href="mxb_print.asp?fromtime=<%=times%>&s=hhz&shopid=<%=defshopvalue%>" title="��ҳ��ʾ�ػ�ױ��">�ػ�ױ��</a> </p>
<table width="100%" border="1" cellspacing="0">
  <tr  bgcolor="#66FFFF" align="center" height="25">
    <td width="100" height="25">�ֶ�����<br>
      ����/ʱ��<br>
      ����</td>
    <td width="100"><%=GetAppellation(3, false)%></td>
    <td width="100"><%=GetAppellation(4, true)%></td>
    <td width="100">��δ��</td>
    <td width="100">&nbsp;����</td>
    <td>��ע</td>
  </tr>
  <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select s.* from shejixiadan s inner join kehu k on s.kehu_id=k.id where s.hhz_time=#"&times&"#"&sqlshop&" order by s.hhz asc",conn,1,1
while not rs.eof
	money1=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
	if isnull(money1) then money1=0
  	money2=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id"))(0)
	if isnull(money2) then money2=0
	money3=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id"))(0)
	if isnull(money3) then money3=0
	money4=conn.execute("select jixiang_money from shejixiadan where id="&rs("id"))(0)
	if isnull(money4) then money4=0
	money5=conn.execute("select sum(money) from save_money where xiangmu_id="&rs("id"))(0)
	if isnull(money5) then money5=0
%>
  <tr align="center" bgcolor="#FFFFFF" onMouseOver="this.bgColor='#FF9966'" onMouseOut="this.bgColor='#FFFFFF'" onClick="javascript:openEditScript('<%=pageurl%>?id=<%=rs("id")%>',450,500)" style="cursor:hand">
    <td  height="25"><%
	if rs("danhao")<>"" and not isnull(rs("danhao")) then
		response.write rs("danhao")&"<br>"
	end if
	%>
        <%=rs("id")%>/<%=rs("hhz")%><br>
        <%=rs("hz_time")%></td>
    <td><%
	arr = GetCustName(rs("kehu_id"))
	if not isarray(arr) then
		redim arr(1)
		arr(0) = "N/A"
		arr(1) = "N/A"
	end if
	response.write arr(0)
	%></td>
    <td><%=arr(1)%></td>
    <td>&nbsp;<%=money1+money2+money3+money4-money5%></td>
    <td>&nbsp;
        <%
	response.write GetFieldDataBySQL("select peplename from yuangong where username='"&rs("userid")&"'","str","")%></td>
    <td><%=rs("yx_cp_memo")%>&nbsp;</td>
  </tr>
  <%
rs.movenext
wend
rs.close
set rs=nothing
%>
</table>
<%end sub%>
<p align="center" class="STYLE4"><br>
  <br>
</p>
<script language="javascript">loadingHidden();</script>
</body>
</html>
