<!--#include file="connstr.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim ProEditFlag
ProEditFlag = true
if session("level")=1 then
	Day_Wed_OutVolume = conn.execute("select Day_Wed_OutVolume from sysconfig")(0)
	if Day_Wed_OutVolume=0 then
		ProEditFlag = false
		if session("zhuguan")=1 then
			ProEditFlag = true
		end if
	else
		ProEditFlag = true
	end if
end if
function chk_user()
	if session("level")=10 or session("level")=7 or session("level")=1 then
		chk_user = true
	else
		chk_user = false
	end if
end function
dim newOrderVerify
newOrderVerify = conn.execute("select newOrderVerify from sysconfig")(0)
if session("level")=10 or (session("level")=1 and session("zhuguan")=1) then newOrderVerify=0
%>
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
<script src="../Js/Calendar.js"></script>
<script language="javascript" src="../inc/func.js" type="text/javascript"></script>
<script language="javascript" src="../js/jixiang_look.js" type="text/javascript"></script>
<link href="../Css/TestDate.css" rel="stylesheet">
<link href="../Css/calendar-blue.css" rel="stylesheet">
<link href="zxcss.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style4 {color: #FF0000}
.div_list_body {
width:98%; 
margin:10px; 
}
.div_list_pro {
width:24%;
float:left;
white-space:nowrap;
}
-->
</style>
<script type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
</head>
<script language="javascript">
function chk()
{
//if(!CheckIsNull(document.form1.jixiang,"请选择套系！")) return false;
if(!CheckIsNumericOrNull(document.form1.money,"请填写套系金额！","套系金额填写错误！"))return false;
if(!CheckIsNumericOrNull(document.form1.savemoney,"请填写收款金额！","收款金额填写错误！"))return false;
var summoney = parseFloat(document.form1.money.value);
var factmoney = parseFloat(document.form1.savemoney.value);
if(summoney<factmoney){
	alert("收款金额不能大于实际金额.");
	return false;
}
if(!CheckIsNull(document.form1.beizhu,"请填写下单具体说明！"))return false;
if(!CheckIsDate(document.form1.pz_time,"请输入正确的拍照日期,格式如:2006-10-1")) return false;
if(!CheckIsDate(document.form1.hz_time,"请输入正确的化妆日期,格式如:2006-10-1")) return false;
if(!CheckIsDate(document.form1.kj_time,"请输入正确的选片日期,格式如:2006-10-1")) return false;
if(!CheckIsDate(document.form1.xg_time,"请输入正确的看版日期,格式如:2006-10-1")) return false;
if(!CheckIsDate(document.form1.qj_time,"请输入正确的取件日期,格式如:2006-10-1")) return false;

	var proflag = false;
	for(var i=0;i<document.form1.check.length;i++){
		if(document.form1.check[i].checked){
			proflag = true;
			if(!CheckIsNumeric2(document.getElementById("sl"+document.form1.check[i].value),"数量不能为空并且只能是数字.")){
				gotoPro(document.getElementById("pronum"+document.form1.check[i].value).value);
				return false;
			}
			if (document.getElementById("p"+document.form1.check[i].value).type=="text")
			{
				if(!CheckIsNumeric2(document.getElementById("p"+document.form1.check[i].value),"P数不能为空并且只能是数字.")){
					gotoPro(document.getElementById("pronum"+document.form1.check[i].value).value);
					return false;
				}
			}
		}
	}

	document.form1.submit();
}
function chk1(flag)
{
//if(!CheckIsNull(document.form1.jixiang2,"请选择套系！")) return false;
if(!CheckIsNull(document.form1.money2,"请填写套系金额！")) return false;
<%
if session("level")=7 or session("level")=1 then
	response.write "if (parseInt(document.form1.money2.value)<parseInt(document.form1.money2.defaultValue)){"&vbcrlf
	response.write "alert('权限不足,您不能调整套系金额.');"&vbcrlf
	response.write "return false;"&vbcrlf
	response.write "}"&vbcrlf
end if
%>
if(!CheckIsNull(document.form1.beizhu,"请填写下单具体说明！"))return false;
if(!CheckIsDate(document.form1.pz_time2,"请输入正确的拍照日期,格式如:2006-10-1")) return false;
if(!CheckIsDate(document.form1.hz_time2,"请输入正确的化妆日期,格式如:2006-10-1")) return false;
if(!CheckIsDate(document.form1.kj_time2,"请输入正确的选片日期,格式如:2006-10-1")) return false;
if(!CheckIsDate(document.form1.xg_time2,"请输入正确的看版日期,格式如:2006-10-1")) return false;
<%if session("level")=1 then%>
if(document.form1.o_qj_time.value!=''){
	if(!CheckIsNull(document.form1.qj_time2,"请输入正确的取件日期,格式如:2006-10-1")) return false;
	if(!CheckIsNull(document.form1.qj,"请输入取件时间！")) return false;
}
<%end if%>
if(flag==0){
	var proflag = false;
	for(var i=0;i<document.form1.check.length;i++){
		if(document.form1.check[i].checked){
			proflag = true;
			if(!CheckIsNumeric2(document.getElementById("sl"+document.form1.check[i].value),"数量不能为空并且只能是数字.")){
				gotoPro(document.getElementById("pronum"+document.form1.check[i].value).value);
				return false;
			}
			if (document.getElementById("p"+document.form1.check[i].value).type=="text")
			{
				if(!CheckIsNumeric2(document.getElementById("p"+document.form1.check[i].value),"P数不能为空并且只能是数字.")){
					gotoPro(document.getElementById("pronum"+document.form1.check[i].value).value);
					return false;
				}
			}
		}
	}
}
	document.form1.submit();
}
function show_prolist(chk){
	tr_el = document.getElementById("div_prolist");
	if(chk.checked){
		tr_el.style.display="";
	}
	else{
		tr_el.style.display="none";
	}
}
</script>　
<body topmargin="0" leftmargin="0">
<div id="div_customer"></div>
<%
adminid=session("adminid")
if adminid="" then
	ue_id=conn.execute("select userid from shejixiadan where id="&request("id"))(0)
	adminid=conn.execute("select id from yuangong where username='"&ue_id&"'")(0)
end if
'response.write keh_id&"--"&ue_id&"--"&request("id")
'response.end
select case request("action")
case "edited"
'if conn.execute("select shejiwancheng from shejixiadan  where id="&request("id")&"")(0)=1 then
'response.Write "<script>alert('对不起,该项目已经完成不能再对其进行修改!');history.go(-1)</ script>"
'Response.End
'end if
if cint(request("qj_flag"))=0 and request("check")<>"" then
	id=split(request("check"),", ")
	for i=lbound(id) to ubound(id)
		if not isnumeric(request("sl"&id(i)&"")) then
			response.Write "<script>alert('数量不能为空并且只能是数字，请检查！');history.go(-1)</script>"
			Response.End
		end if
		sl11=sl11+request("sl"&id(i))&", "
		desc11=desc11 & "|" & trim(request("desc"&id(i)))
		if request("p"&id(i))="" then
			page11=page11&"0, "
		else
			page11=page11&request("p"&id(i))&", "
		end if
	next
	
	if len(sl11)<=2 then
		response.Write "<script>alert('请至少选择一个套系内容，并填写数量！');history.go(-1)</script>"
	else
		sl11=left(sl11,len(sl11)-2)
		page11=left(page11,len(page11)-2)
		desc11=mid(desc11,2)
	end if
end if

if request("pz_time2")<>"" then
	sys=conn.execute("select [CpMaxNum] from sysconfig")(0)
	if isnull(sys) then sys=0
	sy_number=conn.execute("select count(*) from shejixiadan where pz_time=#"&request("pz_time2")&"# and isnull(lc_cp) and id<>"&request("id"))(0)
	if sy_number>=sys and sys<>0 then
		response.Write "<script> alert('摄影当天已达到最高摄影人数,请另选择摄影日期！');history.go(-1) </script>"
		response.end  
	end if
end if
if request("kj_time2")<>"" then
	if conn.execute("select count(*) from shejixiadan where (userid='"&session("userid")&"' or userid2='"&session("userid")&"' or userid3='"&session("userid")&"') and id="&request("id"))(0)<=0 then
		kynum=conn.execute("select kyMaxNum from sysconfig")(0)
		if isnull(kynum) then kynum=0
		ky_number=conn.execute("select count(*) from shejixiadan where kj_time=#"&request("kj_time2")&"# and isnull(lc_cp) and id<>"&request("id"))(0)
		if ky_number>=kynum and kynum<>0 then
			response.Write "<script> alert('选片当天已达到最高选片人数,请另选择选片日期！');history.go(-1) </script>"
			response.end  
		end if
	end if
end if
if request("hz_time2")<>"" and request("hz")="" then
  response.Write "<script>alert('请选择化妆具体时间！');history.go(-1)</script>"
  Response.End
  end if
if request("pz_time2")<>"" and request("pz")="" then
  response.Write "<script>alert('请选择拍照具体时间！');history.go(-1)</script>"
  Response.End
  end if
  if request("pz_time22")<>"" and request("pz2")="" then
  response.Write "<script>alert('请选择拍照2具体时间！');history.go(-1)</script>"
  Response.End
  end if
  if request("hhz_time")<>"" and request("hhz")="" then
  response.Write "<script>alert('请选择回婚妆具体时间！');history.go(-1)</script>"
  Response.End
  end if
   if request("kj_time2")<>"" and request("kj")="" then
  response.Write "<script>alert('请选择选片具体时间！');history.go(-1)</script>"
  Response.End
  end if
   if request("xg_time2")<>"" and request("xg")="" then
  response.Write "<script>alert('请选择看版具体时间！');history.go(-1)</script>"
  Response.End
  end if
 if request("xp2_time2")<>"" and request("xp2")="" then
  response.Write "<script>alert('请选择精修外发具体时间！');history.go(-1)</script>"
  Response.End
  end if
 if request("qj_time2")<>"" and request("qj")="" then
  response.Write "<script>alert('请选择取件具体时间！');history.go(-1)</script>"
  Response.End
  end if
   if not isnumeric(request("money")) then
  response.Write "<script>alert('请填写金额！');history.go(-1)</script>"
  Response.End
  end if
  beizhu11=request("beizhu")
  
    if not isnumeric(request("sl2")) then
  response.Write "<script>alert('精选数量填写错误！');history.go(-1)</script>"
  Response.End
  end if
  if trim(request("danhao2"))<>"" then
	  danhao=conn.execute("select count(*) from shejixiadan where danhao='"&request("danhao2")&"' and id<>"&request("id"))(0)
	  if danhao>0 then
	  	response.Write "<script>alert('该单号已经存在，请检查单号是否错误！');history.go(-1)</script>"
	  	Response.End
	  end if
  end if
  conn.execute("update shejixiadan set wc=null where id="&request("id"))

'写入取件时间
if request("qj_time2")<>"" and request("qj")<>"" then
set ssd=conn.execute("select * from qujian where xiangmu_id="&request("id"))
if not ssd.eof then
conn.execute("update qujian set times=#"&request("qj_time2")&"# where xiangmu_id="&request("id"))
else
conn.execute("insert into qujian (xiangmu_id,userid,times) values("&request("id")&",'"&session("userid")&"',#"&request("qj_time2")&"#)")
end if
end if
'写入取件时间

  set rs2=server.CreateObject("adodb.recordset")
  rs2.open "select * from shejixiadan where id="&request("id"),conn,1,3
  rs2("danhao")=request("danhao")
  rs2("yx_cp_name")=request("yx_cp_name")
  rs2("yx_cp_name2")=request("yx_cp_name2")
  rs2("yx_cp_memo")=request("yx_cp_memo")
  rs2("yx_hz_name")=request("yx_hz_name")
  rs2("yx_hzzl_name")=request("yx_hzzl_name")
  rs2("yx_xg_name")=request("yx_xg_name")
  rs2("yx_ky_name")=request("yx_ky_name")
  rs2("yx_jhz_name")=request("yx_jhz_name")
  rs2("yx_jhlf_name")=request("yx_jhlf_name")
rs2("beizhu")=htmlencode2(beizhu11)
if request("hz_time2")<>"" and request("hz")<>"" then
rs2("hz_time")=request("hz_time2")
rs2("hz")=request("hz")
else
rs2("hz_time")=null
rs2("hz")=null
end if
if request("pz_time2")<>"" and request("pz")<>"" then
rs2("pz_time")=request("pz_time2")
rs2("pz")=request("pz")
else
rs2("pz_time")=null
rs2("pz")=null
end if 
if request("pz_time22")<>"" and request("pz2")<>"" then
	rs2("pz_time2")=request("pz_time22")
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
if request("pzlf_time2")<>"" and request("pzlf")<>"" then
	rs2("pzlf_time")=request("pzlf_time2")
	rs2("pzlf")=request("pzlf")
else
	rs2("pzlf_time")=null
	rs2("pzlf")=null
end if 
if request("jhlf_time2")<>"" and request("jhlf")<>"" then
	rs2("jhlf_time")=request("jhlf_time2")
	rs2("jhlf")=request("jhlf")
else
	rs2("jhlf_time")=null
	rs2("jhlf")=null
end if 
if request("kj_time2")<>"" and request("kj")<>"" then
rs2("kj_time")=request("kj_time2")
rs2("kj")=request("kj")
else
rs2("kj_time")=null
rs2("kj")=null
end if
if request("xg_time2")<>"" and request("xg")<>"" then
rs2("xg_time")=request("xg_time2")
rs2("xg")=request("xg")
else
rs2("xg_time")=null
rs2("xg")=null
end if
if request("xp2_time2")<>"" and request("xp2")<>"" then
rs2("xp2_time")=request("xp2_time2")
rs2("xp2")=request("xp2")
else
rs2("xp2_time")=null
rs2("xp2")=null
end if
if request("qj_time2")<>"" and request("qj")<>"" then
rs2("qj_time")=request("qj_time2")
rs2("qj")=request("qj")
else
rs2("qj_time")=null
rs2("qj")=null
end if
if request("hz_time2")<>"" and request("hz")<>"" then
rs2("hz_time")=request("hz_time2")
rs2("hz")=request("hz")
else
rs2("hz_time")=null
rs2("hz")=null
end if
rs2("jhz_style") = request("jhz_style")
if request("o_hz_time")<>"" then
	if request("hz_time2")<>"" and request("hz")<>"" then
		if cdate(request("hz_time2"))<>cdate(request("o_hz_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","hz_time|hz",request("o_hz_time")&"|"&request("o_hz"),request("hz_time2")&"|"&request("hz"))
			Call EditedTimeSaveToReport(request("id"),e,"hz",request("o_hz_time"),request("hz_time2"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","hz_time|hz",request("o_hz_time")&"|"&request("o_hz"),request("hz_time2")&"|"&request("hz"))
		Call EditedTimeSaveToReport(request("id"),e,"hz",request("o_hz_time"),request("hz_time2"))
	end if
end if
if request("o_pz_time")<>"" then
	if request("pz_time2")<>"" and request("pz")<>"" then
		if cdate(request("pz_time2"))<>cdate(request("o_pz_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","pz_time|pz",request("o_pz_time")&"|"&request("o_pz"),request("pz_time2")&"|"&request("pz"))
			Call EditedTimeSaveToReport(request("id"),e,"pz",request("o_pz_time"),request("pz_time2"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","pz_time|pz",request("o_pz_time")&"|"&request("o_pz"),request("pz_time2")&"|"&request("pz"))
		Call EditedTimeSaveToReport(request("id"),e,"pz",request("o_pz_time"),request("pz_time2"))
	end if
end if
if request("o_pzlf_time")<>"" then
	if request("pzlf_time2")<>"" and request("pzlf")<>"" then
		if cdate(request("pzlf_time2"))<>cdate(request("o_pzlf_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","pzlf_time|pzlf",request("o_pzlf_time")&"|"&request("o_pzlf"),request("pzlf_time2")&"|"&request("pzlf"))
			Call EditedTimeSaveToReport(request("id"),e,"pzlf",request("o_pzlf_time"),request("pzlf_time2"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","pzlf_time|pzlf",request("o_pzlf_time")&"|"&request("o_pzlf"),request("pzlf_time2")&"|"&request("pzlf"))
		Call EditedTimeSaveToReport(request("id"),e,"pzlf",request("o_pzlf_time"),request("pzlf_time2"))
	end if
end if
if request("o_jhlf_time")<>"" then
	if request("jhlf_time2")<>"" and request("jhlf")<>"" then
		if cdate(request("jhlf_time2"))<>cdate(request("o_jhlf_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","jhlf_time|jhlf",request("o_jhlf_time")&"|"&request("o_jhlf"),request("jhlf_time2")&"|"&request("jhlf"))
			Call EditedTimeSaveToReport(request("id"),e,"jhlf",request("o_jhlf_time"),request("jhlf_time2"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","jhlf_time|jhlf",request("o_jhlf_time")&"|"&request("o_jhlf"),request("jhlf_time2")&"|"&request("jhlf"))
		Call EditedTimeSaveToReport(request("id"),e,"jhlf",request("o_jhlf_time"),request("jhlf_time2"))
	end if
end if
if request("o_kj_time")<>"" then
	if request("kj_time2")<>"" and request("kj")<>"" then
		if cdate(request("kj_time2"))<>cdate(request("o_kj_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","kj_time|kj",request("o_kj_time")&"|"&request("o_kj"),request("kj_time2")&"|"&request("kj"))
			Call EditedTimeSaveToReport(request("id"),e,"kj",request("o_kj_time"),request("kj_time2"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","kj_time|kj",request("o_kj_time")&"|"&request("o_kj"),request("kj_time2")&"|"&request("kj"))
		Call EditedTimeSaveToReport(request("id"),e,"kj",request("o_kj_time"),request("kj_time2"))
	end if
end if
if request("o_qj_time")<>"" then
	if request("qj_time2")<>"" and request("qj")<>"" then
		if cdate(request("qj_time2"))<>cdate(request("o_qj_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","qj_time|qj",request("o_qj_time")&"|"&request("o_qj"),request("qj_time2")&"|"&request("qj"))
			Call EditedTimeSaveToReport(request("id"),e,"qj",request("o_qj_time"),request("qj_time2"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","qj_time|qj",request("o_qj_time")&"|"&request("o_qj"),request("qj_time2")&"|"&request("qj"))
		Call EditedTimeSaveToReport(request("id"),e,"qj",request("o_qj_time"),request("qj_time2"))
	end if
end if
if request("o_xg_time")<>"" then
	if request("xg_time2")<>"" and request("xg")<>"" then
		if cdate(request("xg_time2"))<>cdate(request("o_xg_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","xg_time|xg",request("o_xg_time")&"|"&request("o_xg"),request("xg_time2")&"|"&request("xg"))
			Call EditedTimeSaveToReport(request("id"),e,"xg",request("o_xg_time"),request("xg_time2"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","xg_time|xg",request("o_xg_time")&"|"&request("o_xg"),request("xg_time2")&"|"&request("xg"))
		Call EditedTimeSaveToReport(request("id"),e,"xg",request("o_xg_time"),request("xg_time2"))
	end if
end if

if request("o_xp2_time")<>"" then
	if request("xp2_time2")<>"" and request("xp2")<>"" then
		if cdate(request("xp2_time2"))<>cdate(request("o_xp2_time")) then
			e = CheckEvent_Add(request("id"),1,"shejixiadan","xp2_time|xp2",request("o_xp2_time")&"|"&request("o_xp2"),request("xp2_time2")&"|"&request("xp2"))
			Call EditedTimeSaveToReport(request("id"),e,"xp2",request("o_xp2_time"),request("xp2_time2"))
		end if
	else
		e = CheckEvent_Add(request("id"),1,"shejixiadan","xp2_time|xp2",request("o_xp2_time")&"|"&request("o_xp2"),request("xp2_time2")&"|"&request("xp2"))
		Call EditedTimeSaveToReport(request("id"),e,"xp2",request("o_xp2_time"),request("xp2_time2"))
	end if
end if

if request("jhz_style")<>request("o_jhz_style") then
	e = CheckEvent_Add(request("id"),7,"shejixiadan","jhz_style",request("o_jhz_style"),request("jhz_style"))
	Call EditedJhzstyleSaveToReport(request("id"),e,request("o_jhz_style"),request("jhz_style"))
end if
if request("sl_22")<>request("o_sl22") then
	e = CheckEvent_Add(request("id"),8,"shejixiadan","sl2",request("o_sl22"),request("sl_22"))
	Call EditedSl2SaveToReport(request("id"),e,request("o_sl22"),request("sl_22"))
end if
if cint(request("money2"))<>cint(request("old_money2")) and request("old_money2")<>"" and not isnull(request("old_money2")) then
	e = CheckEvent_Add(request("id"),5,"shejixiadan","jixiang_money",request("old_money2"),request("money2"))
	Call EditedMoneySaveToReport(request("id"),e,request("old_money2"),request("money2"))
end if
rs2("jixiang")=request("newjixang")
rs2("jixiang_money")=request("money2")
rs2("stated")=request("stated")
rs2("danhao")=request("danhao2")
rs2("sl2")=request("sl_22")
if cint(request("qj_flag"))=0 and request("check")<>"" then
	rs2("sl")=sl11
	rs2("yunyong")=request("check")
	rs2("pagevol")=page11
	rs2("desc")=desc11

	id3=split(request("check"),", ")
	ttt=split(sl11,", ")
	
	id4=split(request("yunyong"),", ")
	ttt2=split(request("shuliang"),", ")
	
	msg_text = ""
	msg_text_sl = ""
	msg_text_up = ""
	msg_text_dw = ""
	
	for ii=lbound(id4) to ubound(id4)
		set rsyy = conn.execute("select yunyong from yunyong where id="&id4(ii))
		if not rsyy.eof then
			yyname = rsyy("yunyong")
		else
			yyname = "N/A"
		end if
		rsyy.close()
		set rsyy=nothing
		
		exflag = false
		for kk=lbound(id3) to ubound(id3)
			if id4(ii)=id3(kk) then
				exflag = true
				if ttt2(ii)<>ttt(kk) then
					msg_text_sl = msg_text_sl&yyname&" "&ttt(kk)&" 件(原 "&ttt2(ii)&" 件). "
					dvalue = ttt2(ii) - ttt(kk)
					if dvalue<0 then
						conn.execute("update yunyong set sl=sl-"&dvalue&" where id="&id4(ii))
					else
						conn.execute("update yunyong set sl=sl+"&dvalue&" where id="&id4(ii))
					end if
				end if
				exit for
			end if
		next
		if not exflag then
			conn.execute("insert into ProRepList (ProID,RepType,ProVol,Xiangmu_ID,Times,AdminID) values ("&id4(ii)&",0,"&ttt2(ii)&","&request("id")&",#"&now()&"#,"&adminid&")")
			if conn.execute("select [type] from yunyong where id="&id4(ii))(0)=1 then
				conn.execute("update yunyong set sl=sl+"&ttt2(ii)&" where id="&id4(ii)&"")
			end if
			msg_text_dw = msg_text_dw&yyname&" ("&ttt2(ii)&" 件). "
		end if
	next
	for ii=lbound(id3) to ubound(id3)
		set rsyy = conn.execute("select yunyong from yunyong where id="&id3(ii))
		if not rsyy.eof then
			yyname = rsyy("yunyong")
		else
			yyname = "N/A"
		end if
		rsyy.close()
		set rsyy=nothing
		
		exflag = false
		for kk=lbound(id4) to ubound(id4)
			if id3(ii)=id4(kk) then exflag = true
		next
		if not exflag then
			conn.execute("insert into ProRepList (ProID,RepType,ProVol,Xiangmu_ID,Times,AdminID) values ("&id3(ii)&",1,"&ttt(ii)&","&request("id")&",#"&now()&"#,"&adminid&")")
			msg_text_up = msg_text_up&yyname&" ("&ttt(ii)&" 件). "
		end if
	next
	
	if msg_text_sl<>"" or msg_text_up<>"" or msg_text_dw<>"" then
		msg_text = session("username")&" "&" 调整套系产品：<br>"
		if msg_text_sl<>"" then
			msg_text = msg_text&"&nbsp;&nbsp;&nbsp;更换数量："&msg_text_sl&"<br>"
		end if
		if msg_text_up<>"" then
			msg_text = msg_text&"&nbsp;&nbsp;&nbsp;添加产品："&msg_text_up&"<br>"
		end if
		if msg_text_dw<>"" then
			msg_text = msg_text&"&nbsp;&nbsp;&nbsp;移除产品："&msg_text_dw&"<br>"
		end if
		
		fieldname = "yunyong|sl"
		value1 = request("yunyong")&"|"&request("shuliang")
		value2 = request("check")&"|"&sl11
		
		e = CheckEvent_Add(request("id"),3,"shejixiadan",fieldname,value1,value2)
		Call EditedCpvolumeSaveToReport(request("id"),e,msg_text)
	end if
	
	'conn.execute("delete from cuenchu where xiangmu_id="&request("id")&" and type3=1")
'	
'	for ii=lbound(id3) to ubound(id3)
'		if conn.execute("select [type] from yunyong where id="&id3(ii)&"")(0)=1 then
'			conn.execute("update yunyong set sl=sl-"&ttt(ii)&" where id="&id3(ii)&"")
'			conn.execute("insert into cuenchu (xiangmu_id,sp_id,sl,type,type2,type3,beizhu,times) values ("&request("id")&","&id3(ii)&","&ttt(ii)&",2,1,1,'"&htmlencode2(beiz9)&"',#"&now&"#)")
'		end if
'	next
end if

if request("beizhu")="" then
	beiz9="--"
else
	beiz9=request("beizhu")
end if
rs2.update
rs2.close
set rs2=nothing

response.Write "<script>alert('修改下单成功!');location='sjxd.asp?id="&request("id")&"&action=edit'</script>"
  %>
<%case "edit"
newjixiang = request("jixiang")

keh_id=conn.execute("select kehu_id from shejixiadan where id="&request("id"))(0)
ue_id=conn.execute("select userid from shejixiadan where id="&request("id"))(0)
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where id="&request("id")&"",conn,1,1
yunyong11 = rs("yunyong")
sl11 = rs("sl")
desc11 = rs("desc")
pagevol = rs("pagevol")

if newjixiang<>"" and isnumeric(newjixiang) then
	dim rsnewjx
	set rsnewjx=server.createobject("adodb.recordset")
	rsnewjx.open "select * from jixiang where id="&newjixiang,conn,1,1
	if not (rsnewjx.eof and rsnewjx.bof) then
		yunyong11 = rsnewjx("yunyong")
		sl11 = rsnewjx("sl")
		desc11 = null
		pagevol = rsnewjx("pagevol")
	end if
	rsnewjx.close
	set rsnewjx=nothing
end if
	
%><form action="sjxd.asp?action=edited&id=<%=request("id")%>" method="post"  name="form1">
  <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#eeeeee" class="xu_kuan">
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="30" align="right" class="font">套系：</td>
      <td class="font">
      <input name="jixiang2" type="hidden" id="jixiang" value="<%=rs("jixiang")%>">
	  <input name="newjixang" type="hidden" id="newjixang" value="<%
	  if newjixiang<>"" and isnumeric(newjixiang) then
			response.write newjixiang
		else
			response.write rs("jixiang")
		end if
	  %>">
	  <%if session("level")=10 or ProEditFlag or (session("level")=7 and session("zhuguan")=1) then%>
      <select name="xiangmu_list" id="xiangmu_list" onChange="MM_jumpMenu('self',this,0)">
	    <%dim rstypelist,rsxmlist
		set rstypelist = server.createobject("adodb.recordset")
		set rsxmlist = server.createobject("adodb.recordset")
		rstypelist.open "select * from companytype where ishidden=0",conn,1,1
		do while not rstypelist.eof%>
		<optgroup label="<%=rstypelist("companytype")%>">
		<%	rsxmlist.open "select * from jixiang where [type]="&rstypelist("id")&" and ishidden=0 order by px,id",conn,1,1
			do while not rsxmlist.eof%>
        <option value="sjxd.asp?id=<%=request("id")%>&jixiang=<%=rsxmlist("id")%>&action=edit" <%
		if newjixiang<>"" and isnumeric(newjixiang) then
			if cint(newjixiang)=rsxmlist("id") then response.write "selected"
		else
			if rs("jixiang")=rsxmlist("id") then response.write "selected"
		end if
		%>><%=rsxmlist("jixiang")%></option>
			<%	rsxmlist.movenext
			loop
			rsxmlist.close%>
		</optgroup><%
			rstypelist.movenext
		loop
		rstypelist.close
		set rstypelist = nothing
		%>
      </select><%else 
	  	response.write conn.execute("select jixiang from jixiang where id="&rs("jixiang"))(0)
	  end if%></td>
      <td align="right" class="font">套系金额：</td>
      <td class="font"><%
	  if session("level")=10 or ProEditFlag or session("level")=7 then
	  	response.write "<input name='money2' type='text' id='money2' size='13' value='"&rs("jixiang_money")&"'><input name='old_money2' type='hidden' id='old_money2' value='"&rs("jixiang_money")&"'>"
	  else
	  	response.Write rs("jixiang_money")&"<input name='money2' type='hidden' id='money2' value='"&rs("jixiang_money")&"'>"
	  end if
	  %>
      （元）</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="31" align="right" class="font">*摄影日期1：</td>
      <td height="31" colspan="3" class="font"><input name="pz_time2" type="text" maxlength="10" id="pz_time2" size="13" value="<%=rs("pz_time")%>">
          <a onClick="return showCalendar('pz_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
      <input name="pz" type="text" size="3" value="<%=rs("pz")%>">&nbsp;点
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
	  <input name="o_pz_time" type="hidden" id="o_pz_time" value="<%=rs("pz_time")%>">
	 <input name="o_pz" type="hidden" id="o_pz" value="<%=rs("pz")%>"></td>
    </tr>
     <tr align="left" valign="middle" bgcolor="#FFFFFF">
       <td height="31" align="right" bgcolor="#ffffff" class="font">摄影日期2：</td>
       <td colspan="3" bgcolor="#ffffff" class="font"><input name="pz_time22" type="text" maxlength="10" id="pz_time22" size="13" value="<%=rs("pz_time2")%>">
         <a onClick="return showCalendar('pz_time22', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
         <input name="pz2" type="text" size="3" value="<%=rs("pz2")%>">
       &nbsp;点&nbsp;&nbsp;	  &nbsp;&nbsp;摄影/化妆备注
      <input name="yx_cp_memo" type="text" id="yx_cp_memo" size="30" value="<%=rs("yx_cp_memo")%>"></td>
     </tr>
     <tr align="left" valign="middle" bgcolor="#FFFFFF">
       <td height="31" align="right" bgcolor="#ffffff" class="font">拍照礼服：</td>
       <td bgcolor="#ffffff" class="font"><input name="pzlf_time2" type="text" maxlength="10" id="pzlf_time2" size="13" value="<%=rs("pzlf_time")%>" />
           <a onClick="return showCalendar('pzlf_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
           <input name="pzlf" type="text" id="pzlf" value="<%=rs("pzlf")%>" size="3">
         点
         <select name="yx_hzzl_name" id="yx_hzzl_name">
           <option value="">预设<%=GetDutyName(14)%></option>
           <%
		set rscp = server.CreateObject("adodb.recordset")
		rscp.open "select id,peplename from yuangong where [level]=14 and isdisabled=0",conn,1,1
		do while not rscp.eof
			response.write "<option value='"&rscp("peplename")&"'"
			if rscp("peplename")=rs("yx_hzzl_name") then response.write " selected"
			response.write ">"&rscp("peplename")&"</option>"
			rscp.movenext
		loop
		rscp.close
		set rscp = nothing
		%>
         </select>
         <input name="o_pzlf_time" type="hidden" id="o_pzlf_time" value="<%=rs("pzlf_time")%>">
         <input name="o_pzlf" type="hidden" id="o_pzlf" value="<%=rs("pzlf")%>"></td>
       <td align="right" class="font">结婚礼服：</td>
       <td class="font"><input name="jhlf_time2" type="text" id="jhlf_time2" size="13" value="<%=rs("jhlf_time")%>" >
           <a onClick="return showCalendar('jhlf_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
           <input name="jhlf" type="text" size="3" value="<%=rs("jhlf")%>">
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
     <tr align="left" valign="middle" bgcolor="#FFFFFF">
       <td height="31" align="right" class="font">*选片日期：</td>
       <td height="31" class="font"><input name="kj_time2" type="text" id="kj_time2" size="13" value="<%=rs("kj_time")%>" >
         <a onClick="return showCalendar('kj_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
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
       <td align="right" class="font">*结婚化妆：</td>
       <td class="font"><input name="hz_time2" type="text" id="hz_time2" maxlength="10"  size="13" value="<%=rs("hz_time")%>">
           <a onClick="return showCalendar('hz_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
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
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="30" align="right" class="font">*看版日期：</td>
      <td height="30" class="font"><input name="xg_time2" type="text" maxlength="10" id="xg_time2" size="13" value="<%=rs("xg_time")%>">
        <a onClick="return showCalendar('xg_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
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
	  <td height="30" align="right" class="font">配送结婚：</td>
	  <td height="30" class="font"><input name="jhz_style" type="checkbox" id="jhz_style" value="1" <%if instr(rs("jhz_style"),"1")>0 then response.write "checked"%>>
      收费妆&nbsp;
      <input name="jhz_style" type="checkbox" id="jhz_style" value="2" <%if instr(rs("jhz_style"),"2")>0 then response.write "checked"%>>
      免费妆
      <input name="o_jhz_style" type="hidden" id="o_jhz_style" value="<%=rs("jhz_style")%>"></td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="30" align="right" class="font">精修外发：</td>
      <td height="30" class="font"><input name="xp2_time2" type="text" id="xp2_time2" size="13" value="<%=rs("xp2_time")%>">
          <a onClick="return showCalendar('xp2_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="xp2" type="text" id="xp2" value="<%=rs("xp2")%>" size="3">
        点
  <input name="o_xp2_time" type="hidden" id="o_xp2_time" value="<%=rs("xp2_time")%>">
  <input name="o_xp2" type="hidden" id="o_xp2" value="<%=rs("xp2")%>"></td>
      <td height="30" align="right" class="font">回婚妆：</td>
      <td height="30" class="font"><input name="hhz_time" type="text" id="hhz_time" maxlength="10"  size="13" value="<%=rs("hhz_time")%>">
        <a onClick="return showCalendar('hhz_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="hhz" type="text" size="3" value="<%=rs("hhz")%>">
点</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="30" align="right" class="font">取件日期：</td>
      <td height="30" class="font"><input name="qj_time2" type="text" id="qj_time2" size="13" value="<%=rs("qj_time")%>">
          <a onClick="return showCalendar('qj_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="qj" type="text" size="3" value="<%=rs("qj")%>">
        点
  <input name="o_qj_time" type="hidden" id="o_qj_time" value="<%=rs("qj_time")%>">
  <input name="o_qj" type="hidden" id="o_qj" value="<%=rs("qj")%>"></td>
      <td height="30" align="right" class="font">&nbsp;</td>
      <td height="30" class="font">&nbsp;</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="30" colspan="4" class="font">&nbsp;手动单号:
        <input name="danhao2" type="text" id="danhao2" size="10" value="<%=rs("danhao")%>">
&nbsp;        &nbsp;毛片回件情况:
        <input name="stated" type="radio" value="1"  <%if rs("stated")=1 then response.Write "checked"%>>
        正常
        <input type="radio" name="stated" value="2" <%if rs("stated")=2 then response.Write "checked"%>>
        急
        <input type="radio" name="stated" value="3" <%if rs("stated")=3 then response.Write "checked"%>>
        特急&nbsp;&nbsp;&nbsp;拍摄多款选
        <input name="sl_22" type="text" id="sl_22" size="7" value="<%=rs("sl2")%>">
		<input name="o_sl22" type="hidden" id="o_sl22" value="<%=rs("sl2")%>">
        张</td>
    </tr>
        <tr align="left" valign="middle" bgcolor="#FFFFFF">
          <td height="14" colspan="4" class="font" style="padding-left:15px"><div id="div_body" class="div_list_body">
            <div id="div_pro_xx" class="div_list_pro">套系信息正在加载...</div>
          </div></td>
    </tr>
  </table>
	<table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="xu_kuan" style="margin-top:5px; margin-bottom:5px">
      <tr align="left" valign="middle" bgcolor="#FFFFFF">
        <td height="30" colspan="4" class="font">&nbsp;<input name="chk_showpro" type="checkbox" id="chk_showpro" value="yes" onClick="show_prolist(this);">
        显示/修改套系产品</td>
      </tr>
  </table>
	<div id="div_prolist" style="display:none">
	  <div id="div_taoxi">
	    <%
		dim typecounter,prochecked
		typecounter = 0
		set rs2=server.CreateObject("adodb.recordset")
		rs2.open "select * from yunyong_type where ishidden=0 order by px asc",conn,1,1%>
        <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
          <%while not rs2.eof 
		  	typecounter = typecounter + 1
  		zz=zz
  %>
          <tr>
            <td bgcolor="#efefef"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr onClick="javascript:showProList(<%=typecounter%>);" style="cursor:pointer">
                  <td width="25"><img src="../images/+.gif" name="<%="img_jxtype_"&typecounter%>" width="20" height="20" border="0" id="<%="img_jxtype_"&typecounter%>"></td>
                  <td><strong><%=rs2("name")%></strong> </td>
                  <td align="right">&nbsp;</td>
                </tr>
              </table></td>
          </tr>
          <tr id="<%="tr_prolist_"&typecounter%>" style="display:none">
            <td bgcolor="#FFFFFF"><%set rs3=server.CreateObject("adodb.recordset")
	rs3.open "select * from yunyong where type_id="&rs2("id")&" and ishidden=0 order by px",conn,1,1
	if not rs3.eof then
		
		%>
                <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
                  <%
			do while not rs3.eof
		%>
                  <tr onMouseOver="this.bgColor='#FFECFF'" onMouseOut="this.bgColor='#FFFFFF'">
                    <%
			for a=1 to 3
				prochecked=false
				zz=zz+1
				if not rs3.eof then
					i=i-1
					if len(zz)=1 then
						zz="00"&zz
					elseif len(zz)=2 then
						zz="0"&zz
					end if
			%>
                    <td  align=center valign="top" width="30%" id="<%="td_"&zz%>"><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td align="center"><div align="left" style="word-space:nowrap">
                            <input type="checkbox" id="check" name="check" value="<%=rs3("id")%>" <%
		  if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then
		  	namelist=namelist&","&rs3("yunyong")       ''''''''''''''''''''''''''
			typelist=typelist&","&rs3("type")       ''''''''''''''''''''''''''
			xclist=xclist&","&rs3("isxc")       ''''''''''''''''''''''''''
			moneylist=moneylist&","&rs3("money")       ''''''''''''''''''''''''''
			costlist=costlist&","&rs3("in_money")       ''''''''''''''''''''''''''
			counterlist=counterlist&","&typecounter       ''''''''''''''''''''''''''
		  	response.Write "checked"
			prochecked=true
		  end if
		%> onClick="EditProList(this,'<%=zz%>','<%=rs3("yunyong")%>','<%=rs3("id")%>','<%=rs3("isxc")%>','<%=rs3("type")%>','<%=rs3("money")%>','<%=typecounter%>','<%=rs3("in_money")%>')" <%if (not isnull(rs("wc_name")) and not isnull("lc_wc")) or not ProEditFlag then response.write "disabled"%>>
                              <%
		if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then
		  	numlist=numlist&","&zz       ''''''''''''''''''''''''''
		  end if
		response.write "<font color='#cccccc'>"&zz&"</font>&nbsp;-&nbsp;"	
		if rs3("pic")<>"" and not isnull(rs3("pic")) then
			response.write "<a href=""javascript:window.open('yunyong_pic.asp?action=view&sid="&request("id")&"&id="&rs3("id")&"','','fullscreen,scrollbars');void(0);"" title='"&rs3("yunyong")&vbcrlf&"价格"&rs3("money")&"元"&vbcrlf&"点击查看套系图片'><font color=red>"&rs3("yunyong")&"</font></a>"
		else
			response.write "<span title='"&rs3("yunyong")&vbcrlf&"价格"&rs3("money")&"元'>"&rs3("yunyong")&"</span>"
		end if	
		if rs3("type3")=1 then response.write "&nbsp;<font color=#999999>[礼服系列]</font>"
		%></a><a name="<%=zz%>"></a>
                              <input name="<%="pronum"&rs3("id")%>" type="hidden" id="<%="pronum"&rs3("id")%>" value="<%=zz%>">
                          </div></td>
                          <td align="right" style="padding-right:10px"><span id="span_desc<%=rs3("id")%>" class="span_nobr" <%if not prochecked then response.write "style='display:none'"%>><%
		dim tmp_xc
		if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then
			if pagevol<>"" and not isnull(pagevol) then
				tt=split(yunyong11,", ")
				for y=lbound(tt) to ubound(tt)
					if trim(tt(y))=trim(rs3("id")) then
						x=split(pagevol,", ")
						tmp_xc=x(y)
						exit for
					end if
				next
				pagelist=pagelist&","&tmp_xc       ''''''''''''''''''''''''''
			else
				tmp_xc=""
				pagelist=pagelist&",0"       ''''''''''''''''''''''''''
			end if
		else
			tmp_xc=""
		end if
		
		if rs3("isxc")=1 then
		%>
                              <input name="p<%=rs3("id")%>" type="text" id="p<%=rs3("id")%>" size="1" value="<%=tmp_xc%>" onBlur="EditPageVol(this,'<%=zz%>')">
                            P&nbsp;
                            <%
		else%>
                            <input name="p<%=rs3("id")%>" type="hidden" id="p<%=rs3("id")%>" value="<%=tmp_xc%>">
                            &nbsp;
                            <%end if%>
                            <input name="sl<%=rs3("id")%>" type="text" id="sl<%=rs3("id")%>" size="1" value="<% if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then 
		  tt=split(yunyong11,", ")
		  for y=lbound(tt) to ubound(tt)
		  if trim(tt(y))=trim(rs3("id")) then t3=y
		  next 
		  x=split(sl11,", ")
		  if not isnull(desc11) and desc11<>"" then 
			  arr_desc=split(desc11,"|")
			  desc=trim(arr_desc(t3))
		  end if
		  response.Write x(t3)
		  sllist=sllist&","&x(t3)       ''''''''''''''''''''''''''
		  end if%>" onBlur="EditProVol(this,'<%=zz%>')">
                            &nbsp;说明<input type="text" name="<%="desc"&rs3("id")%>"id="<%="desc"&rs3("id")%>" size="4" value="<%=desc%>"></span></td>
                        </tr>
                    </table></td>
                    <%
		  rs3.movenext
		  elseif rs3.eof then
		  response.write"<td height=20 width=30% align=center></td>"
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
  set rs2=nothing
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
  if len(numlist)>2 then
  	numlist=mid(numlist,2)
	namelist=mid(namelist,2)
	sllist=mid(sllist,2)
	pagelist=mid(pagelist,2)
	xclist=mid(xclist,2)
	typelist=mid(typelist,2)
	moneylist=mid(moneylist,2)
	costlist=mid(costlist,2)
	counterlist=mid(counterlist,2)
  end if
  %>
        </table>
	  </div>
	  <input name="inp_yunyong" type="hidden" id="inp_yunyong" value="<%=","&numlist&","%>">
      <input name="inp_name" type="hidden" id="inp_name" value="<%=namelist%>">
	  <input name="inp_sl" type="hidden" id="inp_sl" value="<%=sllist%>">
	  <input name="inp_page" type="hidden" id="inp_page" value="<%=pagelist%>">
	  <input name="inp_xc" type="hidden" id="inp_xc" value="<%=xclist%>">
	  <input name="inp_money" type="hidden" id="inp_money" value="<%=moneylist%>">
	  <input name="inp_cost" type="hidden" id="inp_cost" value="<%=costlist%>">
      <input name="inp_counter" type="hidden" id="inp_counter" value="<%=counterlist%>">
      <input name="pageInvisSetting" type="hidden" id="pageInvisSetting" value="0">
      <input name="newOrderVerify" type="hidden" id="newOrderVerify" value="<%=newOrderVerify%>">
      <input name="inp_typecounter" type="hidden" id="inp_typecounter" value="<%=typecounter%>">
</div>
	<br>
  <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="22%" valign="top"><div align="right">备注说明：
      </div></td>
      <td width="78%"><textarea name="beizhu" cols="70" rows="7" id="beizhu" <% if session("level")<>10 and session("level")<>7 then response.Write("readonly")%>><%=encode2(rs("beizhu"))%>&nbsp;</textarea></td>
    </tr>
  </table>
  <div align="center">  <table width="97%" height="47"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="center">
      <%
	  if not isnull(rs("wc_name")) and not isnull("lc_wc") then 
	  	qj_flag = "1"
	  else
	  	qj_flag = "0"
	  end if
	  %>
      <input name="tijiao" type="button" id="确定" value="确定" onClick="return chk1(<%=qj_flag%>);">

  　<input name="qj_flag" type="hidden" id="qj_flag" value="<%=qj_flag%>">　
    <input name="reset" type="button" id="reset" value="返回" onClick="javascript:history.go(-1)">
    <input name="id" type="hidden" id="id" value="<%=request("id")%>">
    <input name="yunyong" type="hidden" id="yunyong" value="<%=rs("yunyong")%>">
    <input name="shuliang" type="hidden" id="shuliang" value="<%=rs("sl")%>">
</div></td>
    </tr>
  </table>
  </div>
</form>
<script language=javascript>
RefreshCookie("<%=numlist%>","<%=namelist%>","<%=sllist%>","<%=pagelist%>","<%=xclist%>","<%=typelist%>","<%=moneylist%>","<%=counterlist%>","<%=costlist%>");
InitListBody();
</script>
<%rs.close
set rs=nothing%>
<%case else
keh_id=request("id")
ue_id=conn.execute("select userid from kehu where id="&request("id"))(0)
%> 

<form action="xiadan_save.asp?action=save" method="post"  name="form1" onSubmit="return check()">
<table width="97%" height="283" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="xu_kuan">

    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td width="16%" height="30" class="font"><div align="right">另选择下单时间：</div></td>
      <td width="34%" class="font"><input name="times" type="text" id="times" value="<%=date%>" size="13" readonly>        
      <a onClick="return showCalendar('times', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a></td>
      <td colspan="2" class="font">&nbsp;如果没另选择下单时间，默认时间为添加当天日期</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="30" class="font"><div align="right">摄影类型：</div></td>
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
<select name="jixiang" id="select" size="1" ONCHANGE="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}">
 <option value="<%if request("ids")<>"" then 
 response.Write request("ids")
 end if %>"><%if request("ids")<>"" then 
 response.Write conn.execute("select jixiang from jixiang where id="&request("ids")&"")(0)
 else
 response.Write "请选择套系"
 end if %></option>
 
  <%set rs1=server.CreateObject("adodb.recordset")
	  rs1.open "select * from jixiang where ishidden=0 order by [type],px",conn,1,1
	  'ids 为套系里的id
  while not rs1.eof%>
  <option value="sjxd.asp?ids=<%=rs1("id")%>&id=<%=request("id")%>"><%=rs1("jixiang")%></option>
  <%rs1.movenext 
		wend 
		rs1.close
		set rs1=nothing%>
</select></td>
	  <%
	  	if request("ids")<>"" then 
	  jxmoney= conn.execute("select money from jixiang where id="&request("ids")&"")(0)
	  else
	  	jxmoney=0
	  end if
		dim ver,pageInvisSetting
		ver = conn.execute("select [version] from sysconfig")(0)
		pageInvisSetting = conn.execute("select pageInvisSetting from sysconfig")(0)
		if ver="Customer" then 
			response.write "<td colspan=2 width='50%'><input type='hidden' name='money' id='money' value=0><input type='hidden' name='savemoney' id='savemoney' value=0><input type='hidden' name='jxmoney' id='jxmoney' value="&jxmoney&"></td>"
		else
	  %>
      <td width="13%" align="right" class="font">套系金额：</td>
      <td width="37%" class="font"><input name="money" type="text" id="money" size="10" value="<%=jxmoney%>">
      （元）&nbsp;&nbsp; 预收金额：
      <input name="savemoney" type="text" id="savemoney" value="0" size="6">
元</td>
	<%end if%>
    </tr>
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="31" class="font"><div align="right">摄影日期1：</div></td>
      <td height="31" class="font"><input name="pz_time" type="text" maxlength="10" id="txtAwardDate" size="13"/ >
        <a onClick="return showCalendar('pz_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
      <input name="pz" type="text" size="3">
&nbsp;点</td>
      <td align="right" bgcolor="#ffffff" class="font">拍照礼服：</td>
      <td bgcolor="#ffffff" class="font"><input name="pzlf_time" type="text" maxlength="10" id="pzlf_time" size="13" />
	  <a onClick="return showCalendar('pzlf_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
          <input name="pzlf" id="pzlf" type="text" size="3">
  &nbsp;点 (可为空)</td></tr>
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="31" class="font"><div align="right">摄影日期2：</div></td>
      <td height="31" class="font"><input name="pz_time2" type="text" maxlength="10" id="txtAwardDate" size="13"/ >
          <a onClick="return showCalendar('pz_time2', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="pz2" type="text" size="3">
  &nbsp;点</td>
      <td class="font"><div align="right">结婚礼服：</div></td>
      <td class="font"><input name="jhlf_time" type="text" maxlength="10" id="jhlf_time" size="13" >
          <a onClick="return showCalendar('jhlf_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="jhlf" type="text" id="jhlf" size="3">
  &nbsp;点 (可为空)</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="30"class="font"><div align="right">选片日期：</div></td>
      <td height="31" class="font"><input name="kj_time" type="text" maxlength="10" id="txtAwardDate" size="13" >
        <a onClick="return showCalendar('kj_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
      <input name="kj" type="text" size="3">
&nbsp;点</td>
      <td height="30"class="font"><div align="right">结婚化妆：</div></td>
      <td height="30" class="font"><input name="hz_time" type="text" maxlength="10" id="hz_time" size="13" >
          <a onClick="return showCalendar('hz_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="hz" type="text" size="3">
  &nbsp;点</td>
    </tr>
    	<tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="30" class="font"><div align="right">看版日期：</div></td>
      <td height="30" class="font"><input name="xg_time" type="text" id="xg_time" size="13" >
        <a onClick="return showCalendar('xg_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="xg" type="text" size="3">
&nbsp;点</td>
      <td height="30" align="right" class="font">配送结婚：</td>
      <td height="30" class="font"><input name="jhz_style" type="checkbox" id="jhz_style" value="1">
        收费妆&nbsp;
  <input name="jhz_style" type="checkbox" id="jhz_style" value="2">
        免费妆</td>
   	</tr>
        <tr align="left" valign="middle" bgcolor="#ffffff">
          <td height="30" align="right" class="font">取件日期：</td>
          <td height="30" class="font"><input name="qj_time" type="text" id="txtAwardDate" size="13" >
            <a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
            <input name="qj" type="text" size="3">
          &nbsp;点</td>
          <td height="30" align="right" bgcolor="#FFFFFF" class="font">&nbsp;</td>
          <td height="30" bgcolor="#FFFFFF" class="font">&nbsp;</td></tr>
        <tr align="left" valign="middle" bgcolor="#ffffff">
          <td height="30" colspan="4" class="font">&nbsp;&nbsp;手动单号:
            <input name="danhao" type="text" id="danhao" size="8">            
            &nbsp; 毛片回件情况:
            <input name="stated" type="radio" value="1" checked>
正常
<input type="radio" name="stated" value="2">
急
<input type="radio" name="stated" value="3">
特急&nbsp;&nbsp;拍摄多款选
<input name="sl2" type="text" id="sl2" size="7" value="<%if request("ids")<>"" then 
response.Write conn.execute("select sl2 from jixiang where id="&request("ids")&"")(0)
end if%>">
张</td>
        </tr>
		<%if request("ids")<>"" then
		dim rstx_info
		set rstx_info = conn.execute("select yunyong,sl,pagevol from jixiang where id="&request("ids"))
		if not (rstx_info.eof and rstx_info.bof) then
			yunyong11 = rstx_info("yunyong")
			sl11 = rstx_info("sl")
			pagevol = rstx_info("pagevol")
		else
			response.write "<script language=javascript>alert('参数错误,本窗口将自动关闭.');window.close();</script>"
			Response.End
		end if
		rstx_info.close
		set rstx_info = nothing 
		%>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" colspan="4" class="font"><div id="div_body" class="div_list_body">
        <div id="div_pro_xx" class="div_list_pro">套系信息正在加载...</div>
      </div></td>
    </tr>
	<%end if%>
</table>
  <div id="div_prolist">
  <div id="div_taoxi">
    <%
	typecounter = 0
	set rs2=server.CreateObject("adodb.recordset")
	rs2.open "select * from yunyong_type where ishidden=0 order by px asc",conn,1,1%>
    <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
      <%while not rs2.eof 
	  typecounter = typecounter + 1
  zz=zz
  %>
      <tr>
        <td bgcolor="#efefef">&nbsp;&nbsp;<strong><%=rs2("name")%></strong></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%set rs3=server.CreateObject("adodb.recordset")
	rs3.open "select * from yunyong where type_id="&rs2("id")&" and ishidden=0 order by px",conn,1,1
	if not rs3.eof then
		
		%>
            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
              <%
			do while not rs3.eof
		%>
              <tr onMouseOver="this.bgColor='#FFECFF'" onMouseOut="this.bgColor='#FFFFFF'">
                <%
			for a=1 to 3
				prochecked=false
				zz=zz+1
				if not rs3.eof then
					i=i-1
					if len(zz)=1 then
						zz="00"&zz
					elseif len(zz)=2 then
						zz="0"&zz
					end if
			%>
                <td  align=center valign="top" width="30%" id="<%="td_"&zz%>"><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
                    <tr>
                      <td align="center"><div align="left" style="word-space:nowrap">
                          <input type="checkbox" id="check" name="check" value="<%=rs3("id")%>" <%
		  if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then
		  	namelist=namelist&","&rs3("yunyong")       ''''''''''''''''''''''''''
			typelist=typelist&","&rs3("type")       ''''''''''''''''''''''''''
			xclist=xclist&","&rs3("isxc")       ''''''''''''''''''''''''''
			moneylist=moneylist&","&rs3("money")       ''''''''''''''''''''''''''
			costlist=costlist&","&rs3("in_money")       ''''''''''''''''''''''''''
			counterlist=counterlist&","&typecounter       ''''''''''''''''''''''''''
		  	response.Write "checked"
			prochecked=true
		  end if
		%> onClick="EditProList(this,'<%=zz%>','<%=rs3("yunyong")%>','<%=rs3("id")%>','<%=rs3("isxc")%>','<%=rs3("type")%>','<%=rs3("money")%>','<%=typecounter%>','<%=rs3("in_money")%>')">
                          <%
		if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then
		  	numlist=numlist&","&zz       ''''''''''''''''''''''''''
		  end if
		response.write "<font color='#cccccc'>"&zz&"</font>&nbsp;-&nbsp;"	
		if rs3("pic")<>"" and not isnull(rs3("pic")) then
			response.write "<a href=""javascript:window.open('yunyong_pic.asp?action=view&sid="&request("id")&"&id="&rs3("id")&"','','fullscreen,scrollbars');void(0);"" title='"&rs3("yunyong")&vbcrlf&"价格"&rs3("money")&"元"&vbcrlf&"点击查看套系图片'><font color=red>"&rs3("yunyong")&"</font></a>"
		else
			response.write "<span title='"&rs3("yunyong")&vbcrlf&"价格"&rs3("money")&"元'>"&rs3("yunyong")&"</span>"
		end if	
		if rs3("type3")=1 then response.write "&nbsp;<font color=#999999>[礼服系列]</font>"
		%>
                          </a><a name="<%=zz%>"></a>
                          <input name="<%="pronum"&rs3("id")%>" type="hidden" id="<%="pronum"&rs3("id")%>" value="<%=zz%>">
                      </div></td>



                      <td align="right" style="padding-right:10px">
					  
					  
					  
					  
					  
					  <span id="span_desc<%=rs3("id")%>" class="span_nobr" <%if not prochecked then response.write "style='display:none'"%>>
<%
		if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then
			if pagevol<>"" and not isnull(pagevol) then
				tt=split(yunyong11,", ")
				for y=lbound(tt) to ubound(tt)
					if trim(tt(y))=trim(rs3("id")) then
						x=split(pagevol,", ")
						tmp_xc=x(y)
						exit for
					end if
				next
				pagelist=pagelist&","&tmp_xc       ''''''''''''''''''''''''''
			else
				tmp_xc=""
				pagelist=pagelist&",0"       ''''''''''''''''''''''''''
			end if
		else
			tmp_xc=""
		end if
		
		if rs3("isxc")=1 and pageInvisSetting=0 then
		%>
                          <input name="p<%=rs3("id")%>" type="text" id="p<%=rs3("id")%>" size="1" value="<%=tmp_xc%>" onBlur="EditPageVol(this,'<%=zz%>')">
                        P&nbsp;
                        <%
		else%>
                        <input name="p<%=rs3("id")%>" type="hidden" id="p<%=rs3("id")%>" value="<%=tmp_xc%>">
                        &nbsp;
                        <%end if%>



                        <input name="sl<%=rs3("id")%>" type="text" id="sl<%=rs3("id")%>" size="1" value="<% if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then 
		  tt=split(yunyong11,", ")
		  for y=lbound(tt) to ubound(tt)
		  if trim(tt(y))=trim(rs3("id")) then t3=y
		  next 
		  x=split(sl11,", ")
		  response.Write x(t3)
		  sllist=sllist&","&x(t3)       ''''''''''''''''''''''''''
		  end if%>" onBlur="EditProVol(this,'<%=zz%>')">
          &nbsp;说明
		  
		  
		  <input type="text" name="<%="desc"&rs3("id")%>"id="<%="desc"&rs3("id")%>" size="4">
		  
		  
		  </span>
		  
		  













		  </td>
                    </tr>
                </table></td>



                <%
		  rs3.movenext
		  elseif rs3.eof then
		  response.write"<td height=20 width=30% align=center></td>"
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
  set rs2=nothing
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
  if len(numlist)>2 then
  	numlist=mid(numlist,2)
	namelist=mid(namelist,2)
	sllist=mid(sllist,2)
	pagelist=mid(pagelist,2)
	xclist=mid(xclist,2)
	typelist=mid(typelist,2)
	moneylist=mid(moneylist,2)
	costlist=mid(costlist,2)
	counterlist=mid(counterlist,2)
  end if
  %>
    </table>
  </div></div>
  <input name="inp_yunyong" type="hidden" id="inp_yunyong" value="<%=","&numlist&","%>">
  <input name="inp_name" type="hidden" id="inp_name" value="<%=namelist%>">
  <input name="inp_sl" type="hidden" id="inp_sl" value="<%=sllist%>">
  <input name="inp_page" type="hidden" id="inp_page" value="<%=pagelist%>">
  <input name="inp_xc" type="hidden" id="inp_xc" value="<%=xclist%>">
  <input name="inp_money" type="hidden" id="inp_money" value="<%=moneylist%>">
  <input name="inp_cost" type="hidden" id="inp_cost" value="<%=costlist%>">
  <input name="inp_counter" type="hidden" id="inp_counter" value="<%=counterlist%>">
  <input name="pageInvisSetting" type="hidden" id="pageInvisSetting" value="<%=pageInvisSetting%>">
  <input name="newOrderVerify" type="hidden" id="newOrderVerify" value="<%=newOrderVerify%>">
  <input name="inp_typecounter" type="hidden" id="inp_typecounter" value="<%=typecounter%>">
<br>
  <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="22%" valign="top"><div align="right">备注说明： </div></td>
      <td width="78%"><textarea name="beizhu" cols="70" rows="7" id="beizhu"><%if request("ids")<>"" then 
	  response.Write encode2(conn.execute("select beizhu from jixiang where id="&request("ids")&"")(0))
	  end if%>--</textarea></td>
    </tr>
  </table>
  <div align="center">  <table width="97%" height="51"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="center">
      <input name="queding" type="button" id="确定" value="确定" onClick="return chk()">
  　　　　　　　　
    <input name="reset" type="button" id="reset" value="返回" onClick="javascript:history.go(-1)">
    <input name="id" type="hidden" id="id" value="<%=request("id")%>">
</div></td>
    </tr>
  </table>
  </div>
</form>
<script language=javascript>
RefreshCookie("<%=numlist%>","<%=namelist%>","<%=sllist%>","<%=pagelist%>","<%=xclist%>","<%=typelist%>","<%=moneylist%>","<%=counterlist%>","<%=costlist%>");
InitListBody();
</script>
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
	response.write"subcat["&count&"] = new Array('"& trim(rssearch("type"))&"','"& trim(rssearch("jixiang"))&"','sjxd.asp?ids="&trim(rssearch("id"))&"&id="&request("id")&"');"
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
