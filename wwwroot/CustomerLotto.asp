<!--#include file="zlsdk.asp"-->
<!--#include file="connstr.asp"-->
<!--#include file="../inc/sms_class.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/imgInfo.asp"-->
<%
response.write "session('zg_adminid')="&session("zg_adminid")
if request.form("hid_check")="true" then
	'Response.Cookies("LottoMsCheck") = "checked"
	'Response.Cookies("LottoMsUserID") = request.form("zg_msname")
	'Response.Cookies("LottoMsCheck").Expires = DateAdd("d", 1, now())
end if
dim checked
checked = true
'if session("level")<>1 and session("level")<>10 and request.Cookies("LottoMsCheck")<>"checked" then
if session("level")<>1 and session("level")<>10 and session("zg_adminid")="" then
	checked = false
else
	checked = true
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/imgzoom.css" rel="stylesheet" type="text/css">
<link href="../Css/calendar-blue.css" rel="stylesheet" type="text/css">
<link href="zxcss.css" rel="stylesheet" type="text/css">
<script src="../Js/Calendar.js" type="text/javascript"></script>
<script src="../js/imgzoom.js" type="text/javascript"></script>
<STYLE>
<!--
.style2 {
	color: #990066;
	font-size: 18px;
}
.style3 {
	font-size: 16px;
	color: #3300CC;
}
.style4 {font-size: 12px}
.style5 {color: #000000}
.style31 {color: #FF0000}
.div_list_body {
width:90%; 
margin:10px 10px 10px 30px; 
}
.div_list_pro {
width:25%;
float:left;
}
.inp1{font-size:12px;color:#808080;}
html,body {
	background-color:#ffffff;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	height:100%;
	width:100%;
	overflow:hidden;
	text-align:center;
}
-->
</STYLE>
<script language="javascript" src="../inc/ajax.js"></script>
<script language="javascript" src="../inc/func.js"></script>
<script language="javascript" src="../Js/jixiang_look.js"></SCRIPT>
<script language="javascript">
function chks()
{
	if(!CheckIsNull(document.form1.shopid,"请选择连锁店.")) return false;
	if(!CheckIsNull(document.form1.menshi,"请选择门市.")) return false;
	if(!CheckIsNull(document.form1.customerlost,"请选择拍照类型.")) return false;
	if(!CheckIsNull(document.form1.lxpeple,"请填写客户姓名."))　return false;
	if(!CheckIsNull(document.form1.city,"请选择客户所在地区."))　return false;
	if(document.form1.lxpeple2.value!=""){
		if(!CheckIsNull(document.form1.city2,"请选择客户所在地区."))　return false;
	}
	if(!CheckIsDate(document.form1.WeddingDay,"请输入正确的结婚日期,格式如:2006-10-1.")) return false;
	if(!CheckIsPhonenumber(document.form1.telephone,"手机号码不正确,灵通号码请前加区号.")) return false;
	if(!CheckIsPhonenumber(document.form1.telephone2,"手机号码不正确,灵通号码请前加区号.")) return false;
	var now=new Date();
	var year=now.getYear();
	if(!CheckIsShortDate(document.form1.chusheng,"请输入正确的生日日期,格式如:8-8或"+(year-25)+"-8-8.\t")) return false;
	if(!CheckIsShortDate(document.form1.chusheng2,"请输入正确的生日日期,格式如:8-8或"+(year-25)+"-8-8.\t")) return false;
	if(document.form1.hqt_username.value!=""){
		if(!CheckIsNull(document.form1.hqt_password,"请输入婚庆通帐户密码.")) return false;
	}
	document.form1.btn_save.disabled=true;
	document.form1.submit();
	
}
function chkms()
{
	<%if not checked then%>
		bAlert("权限验证",jst_loginhtml,400,220,"true");
		return false;
	<%end if%>
}
function onKeyDown(){  
	//event.ctrlKey && 
	if(event.keyCode==113){
		window.close();
	}
} 
document.onkeydown=onKeyDown;
</script>
<script type="text/javascript">var IMGDIR = '/images';var attackevasive = '0';zoomstatus = parseInt(1);</script>
<%
response.write "<script language=javascript>"&vbcrlf
dim str_loginhtml,rszg
str_loginhtml = str_loginhtml & "var jst_loginhtml=new String("""
str_loginhtml = str_loginhtml & "<table width='90%' border='0' cellspacing='0' cellpadding='0'>"
str_loginhtml = str_loginhtml & "<tr><td height='60'>"
str_loginhtml = str_loginhtml & "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
str_loginhtml = str_loginhtml & "<tr><td>您必须有门市主管访问密码才能访问此页。请在下面选择门市主管并输入密码以便继续。</td></tr></table>"
str_loginhtml = str_loginhtml & "<form id='zgfrm' name='zgfrm' method='post' action='' style='display:inline'>"
str_loginhtml = str_loginhtml & "<fieldset style='padding:5px'><legend>门市主管验证</legend>"
str_loginhtml = str_loginhtml & "选择门市：<select name='zg_msname' id='zg_msname'>"
str_loginhtml = str_loginhtml & "<option value=''>请选择...</option>"
set rszg = server.CreateObject("adodb.recordset")
rszg.open "select * from yuangong where [level]=1 and zhuguan=1 order by username",conn,1,1
do while not rszg.eof
str_loginhtml = str_loginhtml & "<option value='"&rszg("username")&"'>"&rszg("peplename")&"</option>"
rszg.movenext
loop
rszg.close
set rszg = nothing
str_loginhtml = str_loginhtml & "</select>"

str_loginhtml = str_loginhtml & "<br />输入密码：<input type='password' name='zg_mspass' id='zg_mspass' /></fieldset>"
str_loginhtml = str_loginhtml & "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
str_loginhtml = str_loginhtml & "<tr><td height='30'><input type='button' name='zg_btnsend' id='zg_btnsend' value=' 提交 ' style='background-color:#efefef' onClick='javascript:CheckZgInfo();' />&nbsp;"
str_loginhtml = str_loginhtml & "<input type='button' name='zg_btnreset' id='zg_btnreset' value='关闭' onclick='javascript:window.close();' style='background-color:#efefef' /></td>"
str_loginhtml = str_loginhtml & "</tr></table></form>"
str_loginhtml = str_loginhtml & "<div id='div_zgmsg'></div>"
str_loginhtml = str_loginhtml & "</td></tr></table>"
str_loginhtml = str_loginhtml & """);"
response.write str_loginhtml&vbcrlf
response.write "window.onload=chkms;"
response.write "</script>"

%>
<%
dim yunyong11,sl11,soption,arrtype,i,k,rs,rss,rs1,rs2,rs3,pz,kj,qj,cp,xg,hz,id,sl,ii,imgpath
dim FSO,pic,gps,bFlag,DD,PWidth,PHeight,Pp,PXWidth,PXHeight,ImgSize,p1
dim zz,a,namelist,typelist,moneylist,numlist,tt,y,t3,x,sllist,ver

dim action,autosend,rsas
action=request.QueryString("action")

ver = conn.execute("select [version] from sysconfig")(0)
if action="add" then
	'添加客户
	dim lxpeple,telephone,address,address2,home_tel,telephone2,home_tel2,customerbeizhu,count11,savemoney
	dim sy_number,sys,beizhu11,danhao,temp,rsxd,kehu_id,userid2,userid3,xiangmu_id
	dim id3,ttt
	
	lxpeple=request("lxpeple")
	telephone=request("telephone")
	address=request("address")
	address2=request("address2")
	home_tel=request("home_tel")
	telephone2=request("telephone2")
	home_tel2=request("home_tel2")
	
	if telephone<>"" or home_tel<>"" or telephone2<>"" or home_tel2<>"" then 
		dim rscheck,sqlcheck
		if telephone<>"" then sqlcheck=sqlcheck&" or telephone='"&telephone&"' or home_tel='"&telephone&"' or telephone2='"&telephone&"' or home_tel2='"&telephone&"'"
		if home_tel<>"" then sqlcheck=sqlcheck&" or telephone='"&home_tel&"' or home_tel='"&home_tel&"' or telephone2='"&home_tel&"' or home_tel2='"&home_tel&"'"
		if telephone2<>"" then sqlcheck=sqlcheck&" or telephone='"&telephone2&"' or home_tel='"&telephone2&"' or telephone2='"&telephone2&"' or home_tel2='"&telephone2&"'"
		if home_tel2<>"" then sqlcheck=sqlcheck&" or telephone='"&home_tel2&"' or home_tel='"&home_tel2&"' or telephone2='"&home_tel2&"' or home_tel2='"&home_tel2&"'"
		sqlcheck=mid(sqlcheck,5)
		sqlcheck="select * from kehu where ("&sqlcheck&")"
		set rscheck=server.createobject("adodb.recordset")
		rscheck.open sqlcheck,conn,1,1
		if not (rscheck.eof and rscheck.bof) then
			response.write "<script language='javascript'>alert('联系人手机或家庭电话号码重复,单击检测重复可查看具体信息.');history.back();</script>"
			response.end
		end if
		rscheck.close
		set rscheck= nothing
	end if
	
	if trim(request("hqt_username"))<>"" then
		if conn.execute("select count(*) from kehu where hqt_username='"&trim(request("hqt_username"))&"'")(0)>0 then
			response.Write "<script>alert('婚庆通帐户名称已被人使用，请更换后重试！');history.back();</script>"
			Response.End
		end if
	end if
	
	customerbeizhu=htmlencode2(request("customerbeizhu"))
	if customerbeizhu="" or isnull(customerbeizhu) then customerbeizhu="&nbsp;"
	 
	 set rs=server.CreateObject("adodb.recordset")
	 rs.open "select * from kehu",conn,1,3
	 rs.addnew
	 rs("shopid")=request("shopid")
	 rs("CustomerLostType")=request("customerlost")
	 if request("js_id")<>"" and isnumeric(request("js_id")) then rs("js_id")=request("js_id")
	  rs("group")=conn.execute("select [group] from yuangong where username='"&request("menshi")&"'")(0)
	 rs("lxpeple")=lxpeple
	if trim(telephone)<>"" and trim(telephone)<>"灵通号码前加区号" then
	 	rs("telephone")=telephone
	end if
	rs("qq")=trim(request("qq"))
	rs("qq2")=trim(request("qq2"))
	
	 rs("city")=request("city")
	 rs("address")=address
	 if request("lxpeple2")<>"" then
	 rs("lxpeple2")=request("lxpeple2")
	 rs("sex2")=request("sex2")
	 if trim(request("telephone2"))<>"" and trim(request("telephone2"))<>"灵通号码前加区号" then
	 	rs("telephone2")=trim(request("telephone2"))
	 end if
	  rs("city2")=request("city2")
	 rs("address2")=address2
	 rs("post2")=request("post2")
	 if request("chusheng2")<>"" and trim(request("chusheng2"))<>"格式如:8-8"  then 
		rs("chusheng2")=request("chusheng2")
	end if
	 rs("home_tel2")=request("home_tel2")
	 end if
	 rs("post")=request("post")
	 if trim(request("chusheng"))<>"" and trim(request("chusheng"))<>"格式如:8-8" then
		rs("chusheng")=trim(request("chusheng"))
	end if
	
	if trim(request("WeddingDay"))<>"" then
		rs("WeddingDay")=request("WeddingDay")
	end if
	rs("JhDateType")=request("JhDateType")
	
 	rs("ShengriType")=request("ShengriType")
 	rs("ShengriType2")=request("ShengriType2")
	
	 rs("home_tel")=home_tel
	 rs("sex")=request("sex")
	 rs("shuoming")=customerbeizhu
	 rs("userid")=request("menshi")
	 rs("userid2")=request("menshi2")
	 if trim(request("hqt_username"))<>"" and trim(request("hqt_password"))<>"" then
		rs("hqt_username")=trim(request("hqt_username"))
		rs("hqt_password")=MD5_16(trim(request("hqt_password")))
	 end if
	 rs("times")=now()
	 rs("islost")=1
	 rs("pianhao")=request("check2")
	 rs.update
	 temp = rs.bookmark
	 rs.bookmark = temp
	 kehu_id=rs("ID")                '客户ID
	 rs.close()
	 
	 if kehu_id="" or not isnumeric(kehu_id) then
	 	response.Write "<script>alert('客户资料添加失败，请重新操作！');history.back();</script>"
	  	Response.End
	 end if
	 
	 '客户添加完
	
	'接单自动发送短信
	if request.form("chk_autosend")="yes" then
		dim un
		un = conn.execute("select peplename from yuangong where username='"&request("menshi")&"'")(0)
		Call SMSAutoPost("lott",0,kehu_id,un)
	end if
	
'	if err.number>0 then
'		conn.execute("delete from save_money where id="&save_id)
'		response.Write "<script language=javascript>alert('设计下单失败.\n"&err.description&"'.);history.back();</ script>"
'	else
		response.Write "<script language=javascript>"
		response.write "alert('客户资料保存成功!');"
		response.write "location.href='CustomerShootSolution.asp?kid="&kehu_id&"';"
		response.write "</script>"
'	end if
	response.end
end if

%>
<title>婚纱管理系统 -- 客户咨商</title>
</head>
<body>
<%if not checked then
	response.write "<form action='CustomerLotto.asp' method='post' name='form1' style='display:inline'><input name='hid_check' type='hidden' value='true'><input name='btn_save' type='submit' style='display:none'></form>"
else%>
<div id="append_parent"></div><div id="ajaxwaitid"></div>
<div id="div_content" style="height:100%; width:100%; overflow-y:auto">
<form action="CustomerLotto.asp?action=add" method="post" name="form1" style="display:inline" onSubmit="document.getElementById('btn_save').style.filter='gray();';document.getElementById('btn_save').disabled=true;">
<div id="div_blank" style="display:none"></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><%
	dim lottoPic
	lottoPic = conn.execute("select lottoPic from sysconfig")(0)
	if lottoPic="" or isnull(lottoPic) then
		response.write "<img src='../img/lotto.jpg'>"
	else
		if left(lottoPic,4)="http" then
			response.write "<img src='"&lottoPic&"'>"
		else
			response.write "<img src='../upload/"&lottoPic&"'>"
		end if
	end if
	%></td>
  </tr>
</table>
<div id="div_customer" style=""><!-- style="display:none"-->
<div style="width:975px; background-color:#FFFFFF; background-position:right top;  background-repeat:no-repeat">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
    <tr valign="middle"><td colspan="31" height="5"></td></tr>
    <tr valign="middle">
      <td colspan="4" class="font"><%
		dim rsshop
		set rsshop=conn.execute("select * from MultipleShopList order by px")
		if rsshop.eof and rsshop.bof then response.write "<span style='display:none'>"
	  %>&nbsp;连锁店:
        <select name="shopid" id="shopid">
          <option value="">请选择</option>
          <option value="0" <%if rsshop.eof and rsshop.bof then response.write "selected"%>>总店</option>
          <%
		    while not rsshop.eof %>
          <option value="<%=rsshop("id")%>"><%=rsshop("shopname")%></option>
          <%rsshop.movenext
			wend 
			%>
        </select>
        <%
		if rsshop.eof and rsshop.bof then response.write "</span>"
		rsshop.close
		 set rsshop=nothing%>        &nbsp;门市1:<%
		 Call ShowUserSelect("menshi", "1", "username", "请选择", session("userid"), 0)
		 %>
         &nbsp;门市2:<%
		 Call ShowUserSelect("menshi2", "1", "username", "请选择", "", 0)
		 %>
         介绍人:
         <input name="js_name" type="text" id="js_name" size="15" onClick="javascript:openkhwidnow();" readonly />
        <input type="button" name="button" id="button" value="清空" style="width:30px; background-color:eee" onClick="javascript:$E('js_name').value='';$E('js_id').value='';">
        <input type="hidden" name="js_id" id="js_id">
 &nbsp;拍照类型:
        <select name="customerlost" id="customerlost">
          <option value="">请选择</option>
          <%
		  	dim rslosttype
			set rslosttype=conn.execute("select * from customerlosttype order by px")
		    while not rslosttype.eof %>
          <option value="<%=rslosttype("id")%>"><%=rslosttype("title")%></option>
          <%rslosttype.movenext
			wend 
			rslosttype.close
			set rslosttype = nothing
			%>
        </select>        &nbsp;结婚日期:
        <select name="JhDateType" id="JhDateType">
          <option value="0" selected>农历</option>
          <option value="1">公历</option>
        </select>
        <input name="WeddingDay" type="text" id="WeddingDay" size="10" />
        <a onClick="return showCalendar('WeddingDay', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG" /></a>&nbsp;&nbsp;
   <input type="checkbox" name="chk_autosend" id="chk_autosend" value="yes"<%
autosend = GetAutoPostFlag("lott")
select case autosend
	case 1
		response.write " checked"
	case -1
		response.write " disabled title='未配置抽奖短信设置'"
end select
%>>
信息</td>
      </tr>
        <tr align="left" valign="middle">
      <td height="40" colspan="4" class="font"><img src="../images/loot_pianhao_title.gif" width="685" height="30" /></td>
    </tr>
    <tr align="left" valign="top">
      <td height="31" colspan="4"><%set rs3=server.CreateObject("adodb.recordset")
	rs3.open "select * from pianhao where iszs=0 and ishidden=0 order by px,id",conn,1,1
	response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
	while not rs3.eof
	response.write "<tr><td>"
	response.Write "<div style='float:left; width:80px'>&nbsp;&nbsp;"
	set FSO=server.createobject("scripting.filesystemobject")
	pic=server.mappath("../upload/"&rs3("pic"))
	if FSO.FileExists(pic) then
		response.write "<a href=""###zoom"" onClick=""zoom(this, '../upload/"&rs3("pic")&"')"" title='点击预览图片'>"
		response.write trim(rs3("title"))
		response.write "</a>"
	else
		response.write trim(rs3("title"))
	end if
	set FSO=nothing
	response.write ":</div>"
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from pianhao_list where iszs=0 and title_id="&rs3("id")&" and ishidden=0 order by px,id",conn,1,1
	while not rs.eof
		response.write "<div style='float:left; width:70px; white-space:nowrap; padding-right:15px'>"%>
          <input type="checkbox" name="check2" value="<%=rs("id")%>" />
          <%response.Write trim(rs("name"))
	  	response.Write "</div>"
	  rs.movenext
	wend 
	rs.close
	set rs=nothing
	response.Write "</td><tr>"
	rs3.movenext
	wend 
	response.write "</table>"
	rs3.close
	set rs3=nothing%></td>
    </tr>
	</table>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" style="border:dashed 1px #999999">
      <tr align="left" valign="middle">
        <td height="35" colspan="4" class="font"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="45%"><img src="../img/order_05.gif" width="440" height="28" /></td>
              <td width="55%" style="padding-left:110px">&nbsp;</td>
            </tr>
        </table></td>
      </tr>
      <tr align="left" valign="middle">
        <td width="217" height="22" class="font">&nbsp;<%=GetAppellation(1)%>:
          <input name="lxpeple" type="text" id="lxpeple" size="20" />        </td>
        <td width="233" height="22" class="font">&nbsp;客人性别:
          <input name="sex" type="radio" value="男" checked="checked" />
          男
          <input type="radio" name="sex" value="女" />
          女</td>
        <td width="260" height="22" class="font">&nbsp;出生日期:
          <select name="ShengriType" id="ShengriType">
              <option value="0" selected>农历</option>
              <option value="1">公历</option>
            </select>
            <input name="chusheng" type="text" id="chusheng" size="12" class="inp1" onFocus="if (this.value == this.defaultValue) this.value='';" onBlur="if (this.value==''){this.value=this.defaultValue;}else{CheckIsShortDate(this,'请输入正确的生日格式,如:8-8或<%=year(date())-25%>-8-8.\t')}" value="格式如:8-8"></td>
        <td width="260" rowspan="3" valign="bottom" class="font"><input type="image" src="../images/btn_loot_save.gif" name="btn_save" value="保存" style="width:161px; height:69px; border:none; background-color:eeeeee" onClick="return chks()" /></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="22" class="font">&nbsp;个人手机:
          <input name="telephone" type="text" id="telephone" size="15" onKeyUp="value=value.replace(/[^\d]/g,'')"onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
          <input name="btn_checktel1" type="button" id="btn_checktel1" value="检测" onClick="javascript:return CheckCustomerInfo(0,'tel','telephone')"></td>
        <td height="22" class="font">&nbsp;固定电话:
          <input name="home_tel" type="text" id="home_tel" size="15" onKeyUp="value=value.replace(/[^\d]/g,'')"onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
          <input name="btn_checktel2" type="button" id="btn_checktel2" value="检测" onClick="javascript:return CheckCustomerInfo(0,'tel','home_tel')"></td>
        <td height="22" class="font">&nbsp;客人QQ:&nbsp;&nbsp;
          <input name="qq" type="text" id="post" size="20" /></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="22" colspan="2" class="font"><div align="left">&nbsp;客人地址:
          <select name="city" id="city" style="width:100px">
                  <option value="">请选择</option>
                  <%set rs=server.CreateObject("adodb.recordset")
					rs.open "select * from address",conn,1,1
					while not rs.eof %>
                  <option value="<%=rs("address")%>"><%=rs("address")%></option>
                  <%rs.movenext
					wend 
					rs.close
					set rs=nothing%>
                </select>
          &nbsp;
          <input name="address" type="text" id="address" size="40" />
        </div></td>
        <td height="22" class="font">&nbsp;邮政编码:
          <input name="post" type="text" id="post5" size="20" /></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="35" colspan="4" class="font"><img src="../img/order_04.gif" width="440" height="28" /></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" class="font">&nbsp;<%=GetAppellation(2)%>:
          <input name="lxpeple2" type="text" id="lxpeple22" size="20" />
            <div align="right"></div></td>
        <td height="20" class="font">&nbsp;客人性别:
          <input name="sex2" type="radio" value="男" />
          男
          <input name="sex2" type="radio" value="女" checked="checked" />
          女 </td>
        <td height="20" class="font">&nbsp;出生日期:
          <select name="ShengriType2" id="ShengriType2">
              <option value="0" selected>农历</option>
              <option value="1">公历</option>
            </select>
            <input name="chusheng2" type="text" id="chusheng2" size="12" class="inp1" onFocus="if (this.value == this.defaultValue) this.value='';" onBlur="if (this.value==''){this.value=this.defaultValue;}else{CheckIsShortDate(this,'请输入正确的记念日格式,如:8-8或<%=year(date())-25%>-8-8.\t')}" value="格式如:8-8"></td>
        <td rowspan="3" valign="top" class="font"><input type="image" src="../images/btn_loot_close.gif" name="Submit" value="关闭" style="width:161px; height:69px; border:none; background-color:eeeeee" onClick="javascript:window.close();return false;" /></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" class="font">&nbsp;个人手机:
          <input name="telephone2" type="text" id="telephone2" size="15" onKeyUp="value=value.replace(/[^\d]/g,'')"onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
          <input name="btn_checktel3" type="button" id="btn_checktel3" value="检测" onClick="javascript:return CheckCustomerInfo(0,'tel','telephone2')"></td>
        <td height="20" class="font">&nbsp;固定电话:
          <input name="home_tel2" type="text" id="home_tel2" size="15" onKeyUp="value=value.replace(/[^\d]/g,'')"onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
          <input name="btn_checktel4" type="button" id="btn_checktel4" value="检测" onClick="javascript:return CheckCustomerInfo(0,'tel','home_tel2')"></td>
        <td height="20" class="font">&nbsp;客人QQ:&nbsp;&nbsp;
          <input name="qq2" type="text" id="qq2" size="20" /></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" colspan="2" class="font"><div align="right" class="style4">
            <div align="left">&nbsp;客人地址:
              <select name="city2" id="city2" style="width:100px">
                  <option value="">请选择</option>
                  <%set rs=server.CreateObject("adodb.recordset")
					rs.open "select * from address",conn,1,1
					while not rs.eof %>
                  <option value="<%=rs("address")%>"><%=rs("address")%></option>
                  <%rs.movenext
					wend 
					rs.close
					set rs=nothing%>
                </select>
              &nbsp;
              <input name="address2" type="text" id="address2" size="40" />
            </div>
          <div align="left"></div>
        </div></td>
        <td height="20" class="font">&nbsp;邮政编码:
          <input name="post2" type="text" id="post2" size="20" /></td>
      </tr>
      <tr align="left" valign="top">
        <td height="39" colspan="4"><table width="800" height="39" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="60" height="39" valign="top" class="font">&nbsp;备注:</td>
              <td valign="top" class="font"><textarea name="customerbeizhu" cols="95" rows="2" id="customerbeizhu"></textarea></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td height="35" colspan="4"><img src="../img/order_06.gif" width="440" height="28"></td>
      </tr>
      <tr>
        <td height="30">&nbsp;用户名：
          <input name="hqt_username" type="text" id="hqt_username" size="20"></td>
        <td height="30">&nbsp;密码：
          <input name="hqt_password" type="password" id="hqt_password" size="20"></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table>
</div>
<br><br><br>
</div>
</form>
</div>
<%end if%>
</body>
</html>