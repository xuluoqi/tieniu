<!--#include file="admin/ZLSDK.asp"-->
<%dim db,conn,connstr%>
<!--#include file="connstr.asp"-->
<!--#include file="inc/sms_class.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/SystemWorkflow.Class.asp"-->
<%dim CompanyType,xp_reselectflag
CompanyType=GetFieldDataBySQL("select CompanyType from sysconfig","int",0)
xp_reselectflag=false
if session("level")=10 then xp_reselectflag=true%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>影楼管理系统 - 确认选片</title>
<style type="text/css">
<!--
.df {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
}
-->
</style>
<link href="zxcss.css" rel="stylesheet" type="text/css">
<link href="admin/zxcss.css" rel="stylesheet" type="text/css">
<script src="Js/Calendar.js"></script>
<script language="javascript" src="inc/func.js" type="text/javascript"></script>
<script language="javascript" src="js/jixiang_look.js" type="text/javascript"></script>
<link href="Css/calendar-blue.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE2 {color: #FFCC99}
body{padding:8px; margin:0px}
fieldset {margin:5px 0px 5px 0px}
-->
</style>
<script language="javascript">
function chk()
{	
	if(document.form1.xp_name)
		if(!CheckIsNull(document.form1.xp_name,"请选择选片员工！")) return false;
	if(!CheckIsNull(document.form1.jixiang,"请选择套系！")) return false;
	var rl = document.all.xg_opt;
	var xgchecked=false;
    for(var i=0;i<rl.length;i++)
    {
    	if(rl[i].checked){
			if(i==0){
				if(!CheckIsNull(document.form1.qj_time,"请选择取件日期！")) return false;
				if(!CheckIsNull(document.form1.qj,"请输入取件时间！")) return false;
			}
			else if(i==1){
				if(!CheckIsNull(document.form1.xg_time,"请选择看版日期！")) return false;
				if(!CheckIsNull(document.form1.xg,"请输入看版时间！")) return false;
				//if(!CheckIsNull(document.form1.xp2_time,"请选择外发日期！")) return false;
				//if(!CheckIsNull(document.form1.xp2,"请输入外发时间！")) return false;
			}
    		xgchecked=true;
			break;
    	}
    }
	if(!xgchecked){
		alert("请选择看版方式.")
		return false;
	}
	if(document.form1.hz_name){
		if(!CheckIsNull(document.form1.hz_name,"请选择<%=GetDutyName(5)%>！")) return false;
		
		if(document.form1.hz_name2nd){
			if (document.form1.hz_name.value!="" && document.form1.hz_name.value==document.form1.hz_name2nd.value){
				alert("<%=GetDutyName(5)%>1与<%=GetDutyName(5)%>2相同，请更换后重试！")	;
				document.form1.hz_name2nd.focus();
				return false;
			}
		}
	}
	if(document.form1.cp_name){
		if(!CheckIsNull(document.form1.cp_name,"请选择摄影师！")) return false;
	}
	if(document.form1.ts_name){
		if(!CheckIsNull(document.form1.ts_name,"请选择调色师！")) return false;
	}
	
	var proflag = false;
	if(document.form1.pageid){
		if(document.form1.pageid.length){
			for(var i=0;i<document.form1.pageid.length;i++){
				if(!CheckIsNumeric(document.getElementById("p"+document.form1.pageid[i].value),"P数量不能为空并且只能是数字.")) return false;
			}
		}
		else{
			if(!CheckIsNumeric(document.getElementById("p"+document.form1.pageid.value),"P数不能为空并且只能是数字.")) return false;
		}
	}
	
	if(!CheckIsNumericOrNull(document.form1.txt_factmoney,"请填写后期收款金额！","后期收款金额填写错误！"))return false;
	if(!CheckIsDate(document.getElementById("savemoney_time"),"请选择后期收款时间！")) return false;
	var summoney = parseFloat(document.getElementById("lb_summoney").innerHTML);
	var factmoney = parseFloat(document.getElementById("txt_factmoney").value)
	if(summoney<factmoney){
		alert("收款金额不能大于实际金额.");
		return false;
	}
	
	for(var c=1;c<=5;c++){
		if(document.getElementById("cp_wedvol"+c))
			if(!CheckIsNumericOrNull(document.getElementById("cp_wedvol"+c),"请输入正确的照片张数！","请输入正确的照片张数！")) return false;
	}
	
	/*var rbl_qjpro = document.getElementsByName("qjpro_type");
	if(rbl_qjpro!=null){
		for(i=0;i<rbl_qjpro.length;i++)
		{
			if(rbl_qjpro[i].checked){
				if(i==0){
					if(!CheckIsNull(document.form1.qj_time,"请选择取件日期！")) return false;
					if(!CheckIsNull(document.form1.qj,"请输入取件时间！")) return false;
				}
				else if(i==1){
					var inp_qjpro_flag=false;
					var inp_qjpro=document.getElementsByName("inp_qjpro");
					for(c=0;c<inp_qjpro.length;c++){
						if(inp_qjpro.checked){
							inp_qjpro_flag=true
							break;
						}
					}
					if(!inp_qjpro_flag){
						alert("请选择需要分批选项的套系产品。");
						return false;
					}
				}
				break;
			}
		}
	}*/
	
	var flag = false;
	var slt1,slt2;
	for(var i=1;i<=10;i++){
		slt1 = document.getElementById("jixiang"+i);
		
		if(slt1.value!=""){
			for(var k=i+1;k<=10;k++){
				slt2 = document.getElementById("jixiang"+k);
				if(slt2.value!=""){
					if(slt1.value == slt2.value){
						alert("后期项目不能相同.");
						return false;
					}
				}
			}
			if(!CheckIsNumericOrNull(document.getElementById("money"+i),"请填写费用！","费用填写错误！"))　return false;
			if(!CheckIsNumericOrNull(document.getElementById("sl"+i),"请填写数量！","数量填写错误！"))　return false;
			if(document.getElementById("pagevol"+i).type=='text'){
				if(!CheckIsNumericOrNull(document.getElementById("pagevol"+i),"请填写相册P数！","相册P数填写错误！"))　return false;
			}
		}
	}
	
	document.form1.tijiao.disabled=true;
	document.form1.submit();
}
function changeInputType(oldControl,row,inputType){
	var controlParent = document.getElementById("td_page"+row);
	if(inputType=='text'){
		controlParent.innerHTML='P数 ';
	}
	else{
		controlParent.innerHTML='';
	}
	var newControl = document.createElement("input");
    newControl.setAttribute("type",inputType);
	newControl.setAttribute("name",oldControl.name);
	newControl.setAttribute("id",oldControl.id);
	newControl.setAttribute("size","3");
    controlParent.appendChild(newControl);
    //controlParent.removeChild(oldControl);
}
function show_quick_addhq(chk){
	tr_el = document.getElementById("fs_hq");
	chk_money = document.getElementById("chk_getallmoney");
	txtmoney = document.getElementById("txt_factmoney");
	chk_wzsk = document.getElementById("wzsk");
	if(!chk.checked){
		tr_el.style.display="";
	}
	else{
		tr_el.style.display="none";
	}
}
function show_quick_xp(chk){
	tr_el = document.getElementById("fs_xp");
	if(!chk.checked){
		tr_el.style.display="";
	}
	else{
		tr_el.style.display="none";
	}
}
function getSumMoney(){
	var sumMoney = 0
	var slt;
	for(var i=1;i<=10;i++){
		slt = document.getElementById("money"+i);
		if(CheckIsNumericNoMsg(slt)){
			sumMoney+=parseFloat(slt.value);
		}
	}
	document.getElementById("lb_summoney").innerHTML = sumMoney;
}
function chk_xpcheckmode(mode){
	if(mode==0){
		$E("tb_xp_editpro").style.display="";
		$E("tb_xp_mutil").style.display="none";
	}
	else if(mode==1){
		$E("tb_xp_editpro").style.display="none";
		$E("tb_xp_mutil").style.display="";
	}
}
</script>
</head>
<body>
<%
dim clsWorkflow, res, currentWork
set clsWorkflow = new SystemWorkflow
clsWorkflow.DBConnection = conn
clsWorkflow.LoadInstance(false)

dim id,rs,xiangmu_id,lfInvis

select case request("action")
case "edited"

function updateXiangmuFlag(Vid,Vmoney)
	if Vmoney>0 then
		conn.execute("update shejixiadan set ReceivablesFlag=0 where id="&Vid)
	end if
end function

id = request.form("id")
if id="" or not isnumeric(id) then
	response.Write "<script> alert('参数错误，请重新操作！');window.close; </script>"
    response.end
end if

'统计当天日选片量
'dim kynum
'kynum = conn.execute("select count(*) from shejixiadan where datevalue(lc_ky)=#"&date()&"# and not isnull(lc_ky)")(0)
'if isnull(lc_ky) then lc_ky=0
'kymaxnum = conn.execute("select kymaxnum from sysconfig")(0)
'if kymaxnum>0 then
'	if kymaxnum<=kynum then
'		response.write "<script>alert('操作失败，今日选片量已达预设最高日选片量！');history.go(-1)< /script>"
'		Response.End
'	end if
'end if

dim sys,sy_number
if request("pz_time")<>"" and request("pz")<>"" then
sys=conn.execute("select [CpMaxNum] from sysconfig")(0)
if isnull(sys) then sys=0
sy_number=conn.execute("select count(*) from shejixiadan where pz_time=#"&request("pz_time")&"#")(0)
if sy_number>=sys and sys<>0 then
  response.Write "<script> alert('摄影当天已达到最高摄影人数,请另选择摄影日期！');history.go(-1) </script>"
  response.end  
end  if
end if
  
  dim ky_name,ky_name2,userid,userid2
  ky_name = request("xp_name")
  ky_name2 = request("xp_name2")
  if ky_name="" then
  	ky_name = session("username")
	ky_name2=""
  end if
  userid = conn.execute("select username from yuangong where peplename='"&ky_name&"'")(0)
  if ky_name2<>"" then
  	userid2 = conn.execute("select username from yuangong where peplename='"&ky_name2&"'")(0)
  end if
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '修改设计下单信息
  dim rs2
  set rs2=server.CreateObject("adodb.recordset")
  rs2.open "select * from shejixiadan where id="&id,conn,1,3
   
'  '留言
'  if request("beizhu")<>"" then
'	set rs3=server.CreateObject("adodb.recordset")
'	sql3="select * from sjs_baobiao"
'	rs3.open sql3,conn,1,3
'	rs3.addnew
'	rs3("xiangmu_id")=request("id")
'	rs3("baobiao")=HTMLEncode2(request("beizhu"))
'	rs3("times")=now()
'	rs3("userid")=userid
'	rs3("topeple")="所有人"
'	rs3.update
'	rs3.close
'	set rs3=nothing
'  end if	

'调整P数
dim pageid,arr_pgid,pi,pj,txyy,txpg,arr_yy,arr_pg,pg_vol,pgflag
pgflag = false
txyy = rs2("yunyong")
txpg = rs2("pagevol")
arr_yy = split(txyy,", ")
pageid = request.form("pageid")
if pageid<>"" then
	arr_pgid = split(pageid,", ")
	for pi = 0 to ubound(arr_yy)
		for pj = 0 to ubound(arr_pgid)
			if cstr(arr_yy(pi))=cstr(arr_pgid(pj)) then
				pgflag = true
				pg_vol = pg_vol & ", " & request.form("p"&arr_pgid(pj))
				exit for
			end if
		next
		if not pgflag then
			pg_vol = pg_vol & ", 0"
		else
			pgflag = false
		end if
	next
	if pg_vol<>"" then pg_vol=mid(pg_vol,3)
	rs2("pagevol") = pg_vol
end if
	
if CompanyType=0 then
	if request("hz_time")<>"" and request("hz")<>"" then
		rs2("hz_time")=request("hz_time")
		rs2("hz")=request("hz")
	else
		rs2("hz_time")=null
		rs2("hz")=null
	end if
end if
	
	dim arr_qjpro,arr_qjprosl,arr_qjprodesc,arr_qjlist,arr_qjsl,ii,kk,exflag,yyname,msg_text_dw
	if request("qjlist")<>request("inp_qjpro") or request("qjdesc")<>request("inp_qjdesc") then
		arr_qjpro = split(request("inp_qjpro"),", ")
		arr_qjprosl = split(request("inp_qjsl"),", ")
		arr_qjprodesc = split(request("inp_qjdesc"),", ")
		
		arr_qjlist = split(request("qjlist"),", ")
		arr_qjsl = split(request("qjsl"),", ")
		
		msg_text_dw = ""
		for ii=0 to ubound(arr_qjlist)
			yyname = GetFieldDataBySQL("select yunyong from yunyong where id="&arr_qjlist(ii),"str","")
			exflag = false
			for kk=0 to ubound(arr_qjpro)
				if arr_qjlist(ii)=arr_qjpro(kk) then
					exflag = true
					exit for
				end if
			next
			if not exflag then				
				conn.execute("insert into ProRepList (ProID,RepType,ProVol,Xiangmu_ID,Times,AdminID) values ("&arr_qjlist(ii)&",0,"&arr_qjsl(ii)&","&id&",#"&now()&"#,"&session("adminid")&")")
				if conn.execute("select [type] from yunyong where id="&arr_qjlist(ii))(0)=1 then
					conn.execute("update yunyong set sl=sl+"&arr_qjsl(ii)&" where id="&arr_qjlist(ii)&"")
				end if
				msg_text_dw = msg_text_dw&yyname&" ("&arr_qjsl(ii)&" 件). "
			end if
		next
		
		dim n_yunyong,n_sl,n_page,n_protype,n_desc,t_desc
		dim arr_n_yunyong,arr_n_sl,arr_n_desc,arr_n_page,r
		arr_n_yunyong = split(rs2("yunyong"),", ")
		arr_n_sl = split(rs2("sl"),", ")
		if not isnull(rs2("desc")) and rs2("desc")<>"" then arr_n_desc = split(rs2("desc"),"|")
		if pg_vol<>"" then 
			arr_n_page = split(pg_vol,", ")
		else
			arr_n_page = split(rs2("pagevol"),", ")
		end if
		for r = 0 to ubound(arr_n_yunyong)
			n_protype = GetFieldDataBySQL("select [type] from yunyong where id="&arr_n_yunyong(r),"num",0)
			if n_protype=0 then
				n_yunyong = n_yunyong & ", " & arr_n_yunyong(r)
				n_sl = n_sl & ", " & arr_n_sl(r)
				n_page = n_page & ", " & arr_n_page(r)
				if not isnull(rs2("desc")) and rs2("desc")<>"" then
					t_desc = arr_n_desc(r)
				else
					t_desc = ""
				end if	
				n_desc = n_desc & "|" & t_desc
			else
				for ii = 0 to ubound(arr_qjpro)
					if arr_n_yunyong(r)=arr_qjpro(ii) then
						n_yunyong = n_yunyong & ", " & arr_n_yunyong(r)
						n_sl = n_sl & ", " & arr_n_sl(r)
						n_page = n_page & ", " & arr_n_page(r)
						n_desc = n_desc & "|" & arr_qjprodesc(ii)
					end if
				next
			end if
		next
		if n_yunyong<>"" then 
			n_yunyong = mid(n_yunyong,3)
			n_sl = mid(n_sl,3)
			n_page = mid(n_page,3)
			n_desc = mid(n_desc,2)
		end if
		
'		response.write "初始:<br>"
'		response.write "rs2('yunyong')="&rs2("yunyong")&"<br>"
'		response.write "rs2('sl')="&rs2("sl")&"<br>"
'		response.write "rs2('pagevol')="&rs2("pagevol")&"<br>"
'		
'		response.write "提交:<br>"
'		response.write "rs2('yunyong')="&n_yunyong&"<br>"
'		response.write "rs2('sl')="&n_sl&"<br>"
'		response.write "rs2('pagevol')="&n_page&"<br>"
		
		rs2("yunyong") = n_yunyong
		rs2("sl") = n_sl
		rs2("pagevol") = n_page
		rs2("desc") = n_desc
		
		if msg_text_dw<>"" then
			msg_text_dw = session("username")&"&nbsp;调整套系产品：<br>&nbsp;&nbsp;&nbsp;移除产品："&msg_text_dw&"<br>"
		end if
		
		dim fieldname,value1,value2,e
		fieldname = "yunyong|sl"
		value1 = request("qjlist")&"|"&request("qjsl")
		value2 = request("inp_qjpro")&"|"&request("qjsl")
		
		e = CheckEvent_Add(id,3,"shejixiadan",fieldname,value1,value2)
		Call EditedCpvolumeSaveToReport(id,e,msg_text_dw)
	end if

if request("o_hz_time")<>"" and CompanyType=0 then
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
if request("xp2_time")<>"" then
	rs2("xp2_time")=request("xp2_time")
	rs2("xp2")=request("xp2")
else
	rs2("xp2_time")=null
	rs2("xp2")=null
end if
if request("qj_time")<>"" then
	rs2("qj_time")=request("qj_time")
	rs2("qj")=request("qj")
else
	rs2("qj_time")=null
	rs2("qj")=null
end if
if request("qj_time2")<>"" then
	rs2("qj_time2")=request("qj_time2")
	rs2("qj2")=request("qj2")
else
	rs2("qj_time2")=null
	rs2("qj2")=null
end if
if request("xg_time")<>"" then
	rs2("xg_time")=request("xg_time")
	rs2("xg")=request("xg")
else
	rs2("xg_time")=null
	rs2("xg")=null
end if
rs2("xg_opt")=cint(request("xg_opt"))

if request("o_hz_time")<>"" then
	if request("hz_time")<>"" and request("hz")<>"" then
		if cdate(request("hz_time"))<>cdate(request("o_hz_time")) then
			Call EditedTimeSaveToReport(request("id"),"hz",request("o_hz_time"),request("hz_time"))
		end if
	else
		Call EditedTimeSaveToReport(request("id"),"hz",request("o_hz_time"),request("hz_time"))
	end if
end if

dim hstype,hssql,vol,hssl
set hstype=server.createobject("adodb.recordset")
hssql = "select * from hs_signtype order by px asc"
hstype.open hssql,conn,1,1

dim tmpuserid,tmpusername,lvname
lvname = GetUserGroupName(session("userid"))
if request("hzexist")<>"true" then
	rs2("hz_name")=request("hz_name")
	rs2("lc_hz")=now()
	tmpusername = GetFieldDataBySQL("select username from yuangong where peplename='"&request("hz_name")&"'","str","")
	if tmpusername<>"" then
		conn.execute("insert into xiadan (userid,xiangmu_id,type,times) values ('"&tmpusername&"','"&rs2("id")&"',5,#"&now()&"#)")
	end if
	if request("hz_name")<>session("username") then
		conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&request("ID")&",'"&session("userid")&"','[拍照妆]代签名"&lvname&"："&session("username")&"','所有人',#"&now()&"#)")
	end if
	
	tmpuserid = GetFieldDataBySQL("select id from yuangong where peplename='"&request("hz_name")&"'","int",0)
	if not (hstype.eof and hstype.bof) then
		hstype.movefirst
		do while not hstype.eof
			hssl = request.form("hstype_hzs_"& hstype("id"))
			if hssl<>"" and isnumeric(hssl) then
				conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&request("id")&","&tmpuserid&","&hssl&")")
			end if
			hstype.movenext
		loop
	end if
end if
if request("hzzlexist")<>"true" then
	if request("hz_name2")<>"" then
		rs2("hz_name2")=request("hz_name2")
		if request("hz_name2")<>session("username") then
			conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&request("ID")&",'"&session("userid")&"','[拍照妆助理]代签名"&lvname&"："&session("username")&"','所有人',#"&now()&"#)")
		end if
		
		tmpuserid = GetFieldDataBySQL("select id from yuangong where peplename='"&request("hz_name2")&"'","int",0)
		if not (hstype.eof and hstype.bof) then
			hstype.movefirst
			do while not hstype.eof
				hssl = request.form("hstype_hzzl_"& hstype("id"))
				if hssl<>"" and isnumeric(hssl) then
					conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&request("id")&","&tmpuserid&","&hssl&")")
				end if
				hstype.movenext
			loop
		end if
	end if
end if 

if request("cpexist")<>"true" then
	dim cp_wedvol,cp_wedvol2
	cp_wedvol = request("cp_wedvol")
	if cp_wedvol="" or not isnumeric(cp_wedvol) then cp_wedvol=0
	rs2("cp_name")=request("cp_name")
	rs2("cp_wedvol")=cp_wedvol
	rs2("cp_memo")=request("cp_memo")
	rs2("lc_cp")=now()
	if request("cp_name")<>session("username") then
		conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&rs2("ID")&",'"&session("userid")&"','[摄影]代签名"&lvname&"："&session("username")&"','所有人',#"&now()&"#)")
	end if
	
	tmpuserid = GetFieldDataBySQL("select id from yuangong where peplename='"&request("cp_name")&"'","int",0)
	
	if not (hstype.eof and hstype.bof) then
		hstype.movefirst
		do while not hstype.eof
			hssl = request.form("hstype_sys_"& hstype("id"))
			if hssl<>"" and isnumeric(hssl) then
				conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&request("id")&","&tmpuserid&","&hssl&")")
			end if
			hstype.movenext
		loop
	end if
	
	tmpusername = GetFieldDataBySQL("select username from yuangong where peplename='"&request("cp_name")&"'","str","")
	if tmpusername<>"" then
		conn.execute("insert into xiadan (userid,xiangmu_id,type,times) values ('"&tmpusername&"','"&rs2("id")&"',4,#"&now()&"#)")
	end if
	if request("cp_name2")<>"" then
		cp_wedvol2 = request("cp_wedvol2")
		if cp_wedvol2="" or not isnumeric(cp_wedvol2) then cp_wedvol2=0
		rs2("cp_name2")=request("cp_name2")
		rs2("cp_wedvol2")=cp_wedvol2
		rs2("cp_memo2")=request("cp_memo2")
		if request("cp_name2")<>session("username") then
			conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&rs2("ID")&",'"&session("userid")&"','[摄影]代签名"&lvname&"："&session("username")&"','所有人',#"&now()&"#)")
		end if
		
		tmpuserid = GetFieldDataBySQL("select id from yuangong where peplename='"&request("cp_name2")&"'","int",0)
		if not (hstype.eof and hstype.bof) then
			hstype.movefirst
			do while not hstype.eof
				hssl = request.form("hstype_sys2_"& hstype("id"))
				if hssl<>"" and isnumeric(hssl) then
					conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&request("id")&","&tmpuserid&","&hssl&")")
				end if
				hstype.movenext
			loop
		end if
		tmpusername = GetFieldDataBySQL("select username from yuangong where peplename='"&request("cp_name2")&"'","str","")
		if tmpusername<>"" then
			conn.execute("insert into xiadan (userid,xiangmu_id,type,times) values ('"&tmpusername&"','"&rs2("id")&"',4,#"&now()&"#)")
		end if
	end if
else
	if request("cp_name2")<>"" then
		cp_wedvol2 = request("cp_wedvol2")
		if cp_wedvol2="" or not isnumeric(cp_wedvol2) then cp_wedvol2=0
		rs2("cp_name2")=request("cp_name2")
		rs2("cp_wedvol2")=cp_wedvol2
		rs2("cp_memo2")=request("cp_memo2")
		if request("cp_name2")<>session("username") then
			conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&rs2("ID")&",'"&session("userid")&"','[摄影]代签名"&lvname&"："&session("username")&"','所有人',#"&now()&"#)")
		end if
		
		tmpuserid = GetFieldDataBySQL("select id from yuangong where peplename='"&request("cp_name2")&"'","int",0)
		if not (hstype.eof and hstype.bof) then
			hstype.movefirst
			do while not hstype.eof
				hssl = request.form("hstype_sys2_"& hstype("id"))
				if hssl<>"" and isnumeric(hssl) then
					conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&request("id")&","&tmpuserid&","&hssl&")")
				end if
				hstype.movenext
			loop
		end if
		tmpusername = GetFieldDataBySQL("select username from yuangong where peplename='"&request("cp_name2")&"'","str","")
		if tmpusername<>"" then
			conn.execute("insert into xiadan (userid,xiangmu_id,type,times) values ('"&tmpusername&"','"&rs2("id")&"',4,#"&now()&"#)")
		end if
	end if
	lfInvis=conn.execute("select scInvis from sysconfig")(0)
	if lfInvis=1 then
		if request("cp_wedvol")<>"" and isnumeric(request("cp_wedvol")) then
			rs2("cp_wedvol")=request("cp_wedvol")
		end if
		if request("cp_wedvol2")<>"" and isnumeric(request("cp_wedvol2")) then
			rs2("cp_wedvol2")=request("cp_wedvol2")
		end if
		if request("cp_wedvol3")<>"" and isnumeric(request("cp_wedvol3")) then
			rs2("cp_wedvol3")=request("cp_wedvol3")
		end if
		if request("cp_wedvol4")<>"" and isnumeric(request("cp_wedvol4")) then
			rs2("cp_wedvol4")=request("cp_wedvol4")
		end if
		if request("cp_wedvol5")<>"" and isnumeric(request("cp_wedvol5")) then
			rs2("cp_wedvol5")=request("cp_wedvol5")
		end if
	end if
end if
if request("cpzlexist")<>"true" then
	if request("cpzl_name")<>"" then
		rs2("cpzl_name")=request("cpzl_name")
		if request("cpzl_name")<>session("username") then
			conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&rs2("ID")&",'"&session("userid")&"','[摄影助理]代签名"&lvname&"："&session("username")&"','所有人',#"&now()&"#)")
		end if
		
		tmpuserid = GetFieldDataBySQL("select id from yuangong where peplename='"&request("cpzl_name")&"'","int",0)
		
		if not (hstype.eof and hstype.bof) then
			hstype.movefirst
			do while not hstype.eof
				hssl = request.form("hstype_syzl_"& hstype("id"))
				if hssl<>"" and isnumeric(hssl) then
					conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&request("id")&","&tmpuserid&","&hssl&")")
				end if
				hstype.movenext
			loop
		end if
	end if
end if 
hstype.close
set hstype=nothing


if request("tsexist")<>"true" and clsWorkflow.IsExist("ts") then
	rs2("xp_name") = request("ts_name")
	rs2("lc_xp") = now()
	if request("ts_name")<>session("username") then
		conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&request("ID")&",'"&session("userid")&"','[调色]代签名"&lvname&"："&session("username")&"','所有人',#"&now()&"#)")
	end if
end if

rs2("stated")=request("stated")
rs2.update
rs2.close
set rs2=nothing

if err.number>0 then
	err.clear()
	response.Write "<script>alert('确认选片操作失败，请重新操作!');history.go(-1);</script>"
	Response.End
end if

dim kn
kn = ky_name
if ky_name2<>"" then kn = kn & ", " & ky_name2
Call clsWorkflow.SignWork("xp", id, kn, null, true)

'conn.execute("update  shejixiadan set lc_ky=now,ky_name='"&ky_name&"',ky_name2='"&ky_name2&"' where id="&id)
 
if request.form("chk_autosend")="yes" then
	Call SMSAutoPost("ky",id,0,ky_name)
end if
 
' dim settingstr
' settingstr=conn.execute("select mstasksetting from sysconfig")(0)
' if isnull(settingstr) then settingstr=""
' if instr(settingstr,"xp")>0 then
' 	conn.execute("update  shejixiadan set lc_xp2=now,xp2_name='"&ky_name&"' where id="&id)
' end if
' if instr(settingstr,"zd")>0 then
' 	conn.execute("update  shejixiadan set lc_zd=now,zd_name='"&ky_name&"' where id="&id)
' end if
 
 '/////////////////////////////////////////////////////////////////
 '添加后期项目
 dim hq_flag,allmoney,counts9,i,cuenchu_id,yy,allmemo
 dim insertid,temp
 hq_flag = false
 allmoney=0
 counts9=0
 set rs=server.CreateObject("adodb.recordset")
 rs.open "select top 1 * from fujia ",conn,1,3
 for i=1 to 10
	if request("jixiang"&i)<>"" Then
		'if conn.execute("select [type] from yunyong where id="&request("jixiang"&i))(0)=1 then
		'	conn.execute("update yunyong set sl=sl-"&request("sl"&i)&" where id="&request("jixiang"&i))
		'	conn.execute("insert into cuenchu (xiangmu_id,sp_id,sl,type,type2,type3,beizhu,times) values ("&id&","&request("jixiang"&i)&","&request("sl"&i)&",2,1,2,'"&htmlencode2(request("beizhu"&i))&"',#"&now&"#)")
		'	cuenchu_id=conn.execute("select max(id) from cuenchu where xiangmu_id="&id)(0)
		'Else
			cuenchu_id = 0
		'end if
		
		rs.addnew
		rs("cuenchu_id")=cuenchu_id
		rs("sl")=request("sl"&i)
		rs("userid")=userid
		rs("userid2")=userid2
		rs("xiangmu_id")=id
		rs("money")=request("money"&i)
		allmoney=allmoney+request("money"&i)
		rs("jixiang")=request("jixiang"&i)
		
		if request("pagevol"&i)<>"" and isnumeric(request("pagevol"&i)) then
			rs("pagevol")=request("pagevol"&i)
		end if
		
		set yy = conn.execute("select yunyong from yunyong where id="&request("jixiang"&i))
		if not yy.eof then
			counts9=counts9+1
			allmemo = allmemo & "<td>"&yy(0)&"/"&request("sl"&i)&"/"&request("money"&i)&"元&nbsp;&nbsp;"&htmlencode2(request("beizhu"&i))&"</td>"
			if counts9 mod 2 = 0 then allmemo = allmemo &"</tr><tr>"
		end if
		yy.close()
		set yy = nothing
		rs("beizhu")=htmlencode2(request("beizhu"&i))
		rs("times")=now()
		rs.update
		
		temp = rs.bookmark
		rs.bookmark = temp
		insertid=rs("ID")
		
		hq_flag = true
		
		'更新收款状态
		Call updateXiangmuFlag(id,request("money"&i))
	end if
 Next
 
 rs.close
 set rs=nothing
 
 if err.number>0 then
	err.clear()
	response.Write "<script>alert('添加后期操作失败，请转到客户面板重新操作!');window.close();</script>"
	Response.End
 end if
 
 if hq_flag then
	'/////////////////////////////////////////////////////////////////
	'添加后期收款
	dim factmoney
	if request.Form("chk_getallmoney")="yes" then
		factmoney = allmoney
	else
		factmoney = request("txt_factmoney")
	end if
	if allmemo<>"" then
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from save_money",conn,1,3
		rs.addnew
		rs("userid")=userid
		rs("group")=conn.execute("select [group] from yuangong where username='"&userid&"'")(0)
		rs("xiangmu_id")=id
		rs("money")=factmoney
		rs("type")=2
		if request("wzsk")=1 then
			rs("wzsk")=1
		end if
		if request("savemoney_time")<>"" then
			rs("times")=cdate(request("savemoney_time")&" "&time())
		else
			rs("times")=now()
		end if
		allmemo = "<table width=98% align=center><tr>"&allmemo&"</tr></table>"
		rs("beizhu")=allmemo
		rs("orderid")=insertid
		rs.update
		rs.close
		set rs=nothing
		Call FinalMoneySum(id,True)
		 
		if err.number>0 then
			err.clear()
			response.Write "<script>alert('添加收款操作失败，请转到客户面板重新操作!');window.close();</script>"
			Response.End
		end if
	end if
 end if
 
 '/////////////////////////////////////////////////////////////////
 '选片信息编辑
 
dim title,content,sql
title="选片报表确认"
content=request.form("content")
 if content<>"" and not isnull(content) then
	dim flag
	flag=conn.execute("select count(*) from news where danhao="&id)(0)
	Select case flag
	case 0  
	  sql="insert into news (newstitle,newstime,newsmessage,danhao) values ('"&title&"',#"&date()&"#,'"&content&"',"&id&")"
	case 1
	  sql="update  news set newstitle='"&title&"',newstime=#"&date()&"#,newsmessage='"&content&"' where danhao="&id
	end select
	conn.execute(sql)
 end if
 
 '/////////////////////////////////////////////////////////////////
 '完成选片
 response.Write "<script language=javascript>"
 response.write "alert('操作完成，选片单号为"&id&"。');"
 response.write "location.href='admin/lc_baobiao.asp?type=ky&id="&id&"';"
 response.write "</script>"
 response.end
 

case "edit"
dim rssearch,sqlsearch,onecount11
set rssearch=server.createobject("adodb.recordset")
sqlsearch = "select * from yunyong where ishidden=0 order by px asc"
rssearch.open sqlsearch,connstr,1,1
response.write"<script language = ""JavaScript"">"
response.write"var onecount11;"
response.write"onecount11=0;"
response.write"subcat11= new Array();"
count = 0
do while not rssearch.eof 
	response.write"subcat11["&count&"] = new Array('"& trim(rssearch("yunyong"))&" - "&rssearch("money")&"','"&trim(rssearch("id"))&"','"&trim(rssearch("type_id"))&"','"&trim(rssearch("isxc"))&"');"
	count = count + 1
	rssearch.movenext
loop
rssearch.close
response.write"onecount11="&count&";"
response.write"function changelocation(locationid,id)"
response.write"{"
response.write"document.getElementById('jixiang'+id).length = 0;" 
response.write"var locationid=locationid;"
response.write"var i;"
response.write"document.getElementById('jixiang'+id).options[0] = new Option('请选择','');"
response.write"for (i=0;i < onecount11; i++)"
response.write"{"
response.write"if (subcat11[i][2] == locationid)"
response.write"{"
response.write"document.getElementById('jixiang'+id).options[document.getElementById('jixiang'+id).length] = new Option(subcat11[i][0], subcat11[i][1]);"
response.write"}"
response.write"}"
response.write"}"
response.write"function changePageShow(el,yunyongid,id)"
response.write"{" 
response.write"var yunyongid=yunyongid;"
response.write"var i;"
response.write"changeInputType(el,id,'hidden');"
response.write"for (i=0;i < onecount11; i++)"
response.write"{"
response.write"if (subcat11[i][1]==yunyongid && subcat11[i][3]==1)"
response.write"{"
response.write"changeInputType(el,id,'text');"
response.write"}"
response.write"}"
response.write"}"
response.write"</script>"

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shejixiadan where id="&request("id")&"",conn,1,1
dim isczc
isczc = conn.execute("select isczc from jixiang where id="&rs("jixiang"))(0)
%><form action="xiadan2.asp?action=edited" method="post"  name="form1"><fieldset>
    <legend>订单信息</legend>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="xu_kuan">
      <tr align="left" valign="middle">
        <td width="11%" height="20" align="right">&nbsp;套系名称：</td>
        <td width="35%" height="20" valign="middle" class="font"><%=conn.execute("select jixiang from jixiang where id="&rs("jixiang")&"")(0)%>
            <input name="jixiang" type="hidden" id="jixiang" value="<%=rs("jixiang")%>">
        &nbsp;&nbsp;金额 <%=rs("jixiang_money")%> 元&nbsp;&nbsp; 套系 <%=rs("sl2")%> 张</td>
        <td width="11%" height="20" align="right" valign="middle" class="font">选片门市1：</td>
        <td width="43%" height="20" valign="middle" class="font"><select name="xp_name" id="xp_name">
            <option value="">请选择...</option>
            <%
			  dim rss
			  set rss = server.CreateObject("adodb.recordset")
			  rss.open "select * from yuangong where level=1 and isdisabled=0",conn,1,1
			  do while not rss.eof
			  %>
            <option value="<%=rss("peplename")%>" <%if rss("peplename")=rs("ky_name") then response.write "selected"%>><%=rss("peplename")%></option>
            <%
			  rss.movenext
			  loop
			  rss.close
			  %>
          </select>
         &nbsp; 选片门市2：
          <select name="xp_name2" id="xp_name2">
            <option value="">请选择...</option>
            <%
			  set rss = server.CreateObject("adodb.recordset")
			  rss.open "select * from yuangong where level=1 and isdisabled=0",conn,1,1
			  do while not rss.eof
			  %>
            <option value="<%=rss("peplename")%>" <%if rss("peplename")=rs("ky_name2") then response.write "selected"%>><%=rss("peplename")%></option>
            <%
			  rss.movenext
			  loop
			  rss.close
			  %>
        </select></td>
      </tr>
      <%if CompanyType=0 then%>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font"><%=GetWorkName("hz")%>日期：</td>
        <td height="20" class="font"><input name="hz_time" type="text" maxlength="10"  size="13" value="<%=rs("hz_time")%>">
        <a onClick="return showCalendar('hz_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG3" /></a>
        <input name="hz" type="text" size="3" value="<%=rs("hz")%>">
        点
        <input name="o_hz_time" type="hidden" id="o_hz_time" value="<%=rs("hz_time")%>">
<input name="o_hz" type="hidden" id="o_hz" value="<%=rs("hz")%>"></td>
        <td height="20" align="right" class="font">&nbsp;</td>
        <td height="20" class="font">&nbsp;</td>
      </tr>
      <%end if%>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font">看版日期：</td>
        <td height="20" class="font"><input name="xg_time" type="text" id="xg_time" size="13"  value="<%=rs("xg_time")%>">
        <a onClick="return showCalendar('xg_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG4" /></a>
        <input name="xg" type="text" size="3" value="<%if rs("xg")="" or isnull(rs("xg")) then
		  	response.write "0"
		 else
		 	response.write rs("xg")
		 end if%>">
点
<input name="o_xg_time" type="hidden" id="o_xg_time" value="<%=rs("xg_time")%>">
<input name="o_xg" type="hidden" id="o_xg" value="<%=rs("xg")%>"></td>
        <td height="20" align="right"><span class="font">看版：</span></td>
        <td height="20"><span class="font">
          <input name="xg_opt" type="radio" value="0">
内部看版
<input type="radio" name="xg_opt" value="1">
客户</span></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font">外发时间：</td>
        <td height="20" class="font"><input name="xp2_time" type="text" id="xp2_time" size="13"  value="<%=rs("xp2_time")%>">
          <a onClick="return showCalendar('xp2_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG6" /></a>
          <input name="xp2" type="text" size="3" value="<%if rs("xp2")="" or isnull(rs("xp2")) then
		  	response.write "0"
		 else
		 	response.write rs("xp2")
		 end if%>">
点
<input name="o_xp2_time" type="hidden" id="o_xp2_time" value="<%=rs("xp2_time")%>">
<input name="o_xp2" type="hidden" id="o_xp2" value="<%=rs("xp2")%>"></td>
        <td height="20" align="right" class="font">毛片回件：</td>
        <td height="20" class="font"><input name="stated" type="radio" value="1"  <%if rs("stated")=1 then response.Write "checked"%>>
正常
  <input type="radio" name="stated" value="2" <%if rs("stated")=2 then response.Write "checked"%>>
急
<input type="radio" name="stated" value="3" <%if rs("stated")=3 then response.Write "checked"%>>
特急 </td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font">取件时间1：</td>
        <td height="20" class="font"><input name="qj_time" type="text" id="qj_time" size="13"  value="<%=rs("qj_time")%>">
          <a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG" /></a>
          <input name="qj" type="text" size="3" value="<%if rs("qj")="" or isnull(rs("qj")) then
		  	response.write "0"
		 else
		 	response.write rs("qj")
		 end if%>">
点
<input name="o_qj_time" type="hidden" id="o_qj_time" value="<%=rs("qj_time")%>">
<input name="o_qj" type="hidden" id="o_qj" value="<%=rs("qj")%>"></td>
        <td height="20" align="right" class="font">&nbsp;</td>
        <td height="20" class="font">&nbsp;</td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font">取件时间2：</td>
        <td height="20" class="font"><input name="qj_time2" type="text" id="qj_time2" size="13"  value="<%=rs("qj_time2")%>">
          <a onClick="return showCalendar('qj_time2', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG" /></a>
          <input name="qj2" type="text" id="qj2" value="<%if rs("qj2")="" or isnull(rs("qj2")) then
		  	response.write "0"
		 else
		 	response.write rs("qj2")
		 end if%>" size="3">
          点
  <input name="o_qj_time2" type="hidden" id="o_qj_time2" value="<%=rs("qj_time2")%>">
  <input name="o_qj2" type="hidden" id="o_qj2" value="<%=rs("qj2")%>"></td>
        <td height="20" class="font">&nbsp;</td>
        <td height="20" class="font">&nbsp;</td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font"><%=GetDutyName(5)%>1：</td>
        <td height="20" colspan="3" class="font"><%
		if not isnull(rs("hz_name")) and rs("hz_name")<>"" then
			Call ShowUserSelect("hz_name", "5, 14", "peplename", "请选择...", rs("hz_name"), 100, true)
			response.write "<input type='hidden' id='hzexist' name='hzexist' value='true'>"
		else
			Call ShowUserSelect("hz_name", "5, 14", "peplename", "请选择...", rs("hz_name"), 100, false)
		end if
		response.write "&nbsp;&nbsp;"
		response.write ShowWedSignInput("hstype_hzs_", rs("id"), rs("hz_name"), true)
		%></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font"><%=GetDutyName(5)%>2：</td>
        <td height="20" colspan="3" class="font"><%
		if not isnull(rs("hz_name2nd")) and rs("hz_name2nd")<>"" then
			Call ShowUserSelect("hz_name2nd", "5, 14", "peplename", "请选择...", rs("hz_name2nd"), 100, true)
			response.write "<input type='hidden' id='hz2exist' name='hz2exist' value='true'>"
		else
			Call ShowUserSelect("hz_name2nd", "5, 14", "peplename", "请选择...", rs("hz_name2nd"), 100, false)
		end if
		response.write "&nbsp;&nbsp;"
		response.write ShowWedSignInput("hstype_hzs2_", rs("id"), rs("hz_name2nd"), true)
		%></td>
      </tr>
      <tr align="left" valign="middle">
        <td height="20" align="right" valign="top" class="font"><%=GetDutyName(14)%>：</td>
        <td height="20" colspan="3" class="font"><%
		if not isnull(rs("hz_name2")) and rs("hz_name2")<>"" then
			Call ShowUserSelect("hz_name2", "14", "peplename", "请选择...", rs("hz_name2"), 100, true)
			response.write "<input type='hidden' id='hzzlexist' name='hzzlexist' value='true'>"
		else
			Call ShowUserSelect("hz_name2", "14", "peplename", "请选择...", rs("hz_name2"), 100, false)
		end if
		response.write "&nbsp;&nbsp;"
		response.write ShowWedSignInput("hstype_hzzl_", rs("id"), rs("hz_name2"), true)
		%></td>
      </tr>
	  <%if isnull(rs("cp_name")) or rs("cp_name")="" then%>
      <tr align="left" valign="middle">
        <td height="20" align="right" valign="top" class="font">摄影师：</td>
        <td height="20" colspan="3" class="font"><%
		Call ShowUserSelect("cp_name", "4, 12", "peplename", "请选择...", rs("cp_name"), 100, false)%>
        <!--&nbsp;拍摄服装：-->
        <%'response.write "<input type='text' id='cp_memo' name='cp_memo' value=''>"
		lfInvis=conn.execute("select scInvis from sysconfig")(0)
		if lfInvis=1 then
		%>
&nbsp; 入选总片
<input name="cp_wedvol" type="text" id="cp_wedvol" size="5">
张<%end if
		response.write "&nbsp;&nbsp;"
		response.write ShowWedSignInput("hstype_sys_", rs("id"), rs("cp_name"), true)
%><br /><%
		Call ShowUserSelect("cp_name2", "4, 12", "peplename", "请选择...", rs("cp_name2"), 100, false)%>
        <!--&nbsp;拍摄服装：-->
        <%'response.write "<input type='text' id='cp_memo2' name='cp_memo2' value=''>"
		lfInvis=conn.execute("select scInvis from sysconfig")(0)
		if lfInvis=1 then
		%>
&nbsp; 入选总片
<input name="cp_wedvol" type="text" id="cp_wedvol" size="5">
张<%end if

	  	response.write "&nbsp;"
		response.write ShowWedSignInput("hstype_sys2_", rs("id"), rs("cp_name2"), true)
%></td>
      </tr>
	  <%else%>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font" valign="top">摄影师：</td>
        <td height="20" colspan="3" class="font">
        <%
		response.write "<input type='hidden' id='cpexist' name='cpexist' value='true'>"
		dim k,flag1,cp_id,str1
		for k = 1 to 5
			flag1 = false
			if k = 1 then 
				cp_id = ""
			else
				cp_id = k
			end if
			if (rs("cp_name"&cp_id)<>"" and not isnull(rs("cp_name"&cp_id))) or k=2 then
				flag1 = true
				if rs("cp_name"&cp_id)<>"" and not isnull(rs("cp_name"&cp_id)) then 
					str1=true
				else
					str1=false
				end if
				Call ShowUserSelect("cp_name"&cp_id, "4", "peplename", "请选择...", rs("cp_name"&cp_id), 100, str1)%>
				<!--&nbsp;拍摄服装：-->
                <%'if rs("cp_memo"&cp_id)="" or isnull(rs("cp_memo"&cp_id)) then
					'response.write "无"
				  'else
				  '	response.write rs("cp_memo"&cp_id)
				  'end if
				  lfInvis=conn.execute("select scInvis from sysconfig")(0)
				  if lfInvis=1 then
		%>
               &nbsp; 入选总片
				<input name="<%="cp_wedvol"&cp_id%>" type="text" id="<%="cp_wedvol"&cp_id%>" size="5" value="<%=rs("cp_wedvol"&cp_id)%>">
			张<%
				end if
				response.write "&nbsp;"
				response.write ShowWedSignInput("hstype_sys"&cp_id&"_", rs("id"), rs("cp_name"&cp_id), true)
				
				if flag1 then response.write "<br>"
			end if
		next
		%></td>
      </tr>
	  <%end if%>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font">摄影助理：</td>
        <td height="20" colspan="3" class="font"><%
		if not isnull(rs("cpzl_name")) and rs("cpzl_name")<>"" then
			Call ShowUserSelect("cpzl_name", "12", "peplename", "请选择...", rs("cpzl_name"), 100, true)
			response.write "<input type='hidden' id='cpzlexist' name='cpzlexist' value='true'>"
		else
			Call ShowUserSelect("cpzl_name", "12", "peplename", "请选择...", rs("cpzl_name"), 100, false)
		end if
		
		response.write "&nbsp;&nbsp;"
		response.write ShowWedSignInput("hstype_syzl_", rs("id"), rs("cpzl_name"), true)
		%></td>
      </tr>
      <%if clsWorkflow.IsExist("ts") then%>
      <tr align="left" valign="middle">
        <td height="20" align="right" class="font">调色：</td>
        <td height="20" colspan="3" class="font"><%
		if not isnull(rs("xp_name")) and rs("xp_name")<>"" then
			Call ShowUserSelect("ts_name", "2,4,12", "peplename", "请选择...", rs("xp_name"), 100, true)
			response.write "<input type='hidden' id='tsexist' name='tsexist' value='true'>"
		else
			Call ShowUserSelect("ts_name", "2,4,12", "peplename", "请选择...", rs("xp_name"), 100, false)
		end if
		%></td>
      </tr>
      <%end if%>
      </table>
</fieldset>
  <fieldset>
  <legend>自动</legend>
  <div style="width:98%; padding:5px"><span class="font">
    <input type="checkbox" name="chk_autosend" id="chk_autosend" value="yes"<%
dim autosend
autosend = GetAutoPostFlag("ky")
select case autosend
	case 1
		response.write " checked"
	case -1
		response.write " disabled title='未配置选片短信设置'"
end select
%>>
    信息</span></div>
  </fieldset>
  <fieldset>
  <legend>相册P数调整</legend>
  <%
  dim idlist,sllist,wclist,desclist
  if isnull(rs("yunyong")) then
		response.Write "<br>没有套系应有!"
	else
		idlist=split(rs("yunyong"),", ")
		
		if not isnull(rs("wc")) then
			wclist=split(rs("wc"),", ")
		end if
		if not isnull(rs("desc")) and rs("desc")<>"" then 
			desclist=split(rs("desc"),"|")
		end if
%>
  <div style="width:98%; padding:5px">
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <%
	  dim count11,count22,rslistflag
	  count11=ubound(idlist)+1
	  if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			sllist=split(rs("pagevol"),", ")
	  end if
	  count22=0
	  for yy=1 to count11
		
		set rslistflag = conn.execute("select [isxc] from yunyong where id="&idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("isxc")=1 then
				count22=count22+1
	%>
        <td><%
		response.write "<table width='85%'  border='0' cellspacing='0' cellpadding='0'><tr><td>"
		if len(count22)=2 then
			response.Write "<strong>"&count22&".</strong>"
		else
			response.Write "<strong>0"&count22&"</strong>"
			response.Write "."
		end if
		
		dim yyflag,rslist_yunyong
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		response.Write "<input type='hidden' id='pageid' name='pageid' value='"&idlist(yy-1)&"'>"
		response.Write "<input type='text' id='p"&idlist(yy-1)&"' name='p"&idlist(yy-1)&"' value='"
		if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			response.Write sllist(yy-1)
		end if
		response.write "' size='3'> P"
		rslist_yunyong.close()
		response.write "</td></tr></table>"
		%></td>
        <%
				if count22 mod 3 =0 then response.write "</tr><tr>"
			end if
			end if
			rslistflag.close()
		next
		%>
       </tr>
    </table>
  </div>
  <%end if%>
  </fieldset>
  <fieldset>
<legend>调整取件内容</legend>
  <%
  if isnull(rs("yunyong")) then
		response.Write "<br>没有套系应有!"
	else
		dim qjlist,qjsl,qjdesc,tmpdesc
		qjlist=""
		qjsl=""
		qjdesc=""
		idlist=split(rs("yunyong"),", ")
		sllist=split(rs("sl"),", ")
		if not isnull(rs("wc")) then
			wclist=split(rs("wc"),", ")
		end if
%>
<div style="width:98%; padding:5px">
  <table id="tb_xp_editpro" width="100%"  border="0" cellspacing="0" cellpadding="0">
	<tr><%
	  count11=ubound(idlist)+1
	  count22=0
	  for yy=1 to count11
		set rslistflag = conn.execute("select * from yunyong where id="&idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("type")=1 then
				tmpdesc=""
				count22=count22+1
				qjlist = qjlist&", "&idlist(yy-1)
				qjsl = qjsl&", "&sllist(yy-1)%>
                <td width="33.3%"><%
                response.write "<div style='float:left; width:69.999%;'>"
                response.Write "<input type='checkbox' name='inp_qjpro' id='inp_qjpro' value='"&idlist(yy-1)&"' checked>"
                response.Write "<input type='hidden' name='inp_qjsl' id='inp_qjsl' value='"&sllist(yy-1)&"'>&nbsp;"
                response.Write rslistflag("yunyong")
                response.Write " - "&sllist(yy-1)
                response.write "</div><div style='float:left; text-align:right; width:30%;'>"
                response.write "&nbsp;说明&nbsp;<input type='text' name='inp_qjdesc' id='inp_qjdesc' size='4' value="""
                if not isnull(rs("desc")) and rs("desc")<>"" then 
                	if desclist(yy-1)<>"" then tmpdesc = desclist(yy-1)
                end if
                qjdesc=qjdesc & ", " & tmpdesc
                response.write tmpdesc
                response.write """>&nbsp;</div>"
                %></td>
		<%		if count22 mod 3 =0 then response.write "</tr><tr>"
			end if
		end if
		rslistflag.close()
	  next
	  if qjlist<>"" then qjlist = mid(qjlist,3)
	  if qjsl<>"" then qjsl = mid(qjsl,3)
	  if qjdesc<>"" then qjdesc=mid(qjdesc,3)
	  %>
		<input type="hidden" name="qjlist" id="qjlist" value="<%=qjlist%>">
		<input type="hidden" name="qjsl" id="qjsl" value="<%=qjsl%>">
		<input type="hidden" name="qjdesc" id="qjdesc" value="<%=qjdesc%>">
        </tr>
  </table>
</div>
<%end if%></fieldset>
  <fieldset>
    <legend>分批选片</legend>
    <%
  if isnull(rs("yunyong")) then
		response.Write "<br>没有套系应有!"
	else
		qjlist=""
		qjsl=""
		qjdesc=""
		idlist=split(rs("yunyong"),", ")
		sllist=split(rs("sl"),", ")
		if not isnull(rs("wc")) then
			wclist=split(rs("wc"),", ")
		end if
%>
    <div style="width:98%; padding:5px">
      <table id="tb_xp_mutil" width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <%
	  dim xp_qjlist,xp_qjsl,yx_flag
	  count11=ubound(idlist)+1
	  count22=0
	  for yy=1 to count11
		set rslistflag = conn.execute("select * from yunyong where id="&idlist(yy-1))
		if not rslistflag.eof then
			if rslistflag("type")=1 then
				count22=count22+1
				yx_flag=getfielddatabysql("SELECT flow_xp_proinfo.id FROM flow_xp_prolist INNER JOIN flow_xp_proinfo ON flow_xp_prolist.id = flow_xp_proinfo.listid where flow_xp_prolist.xiangmu_id="&request("id")&" and flow_xp_proinfo.proid="&idlist(yy-1),"int",0)%>
          <td width="33.3%"><%
                'response.write "<div style='float:left; width:69.999%;'>"
                response.Write "<input type='checkbox' name='inp_xp_qjpro' id='inp_xp_qjpro' value='"&idlist(yy-1)&"'"
				if yx_flag<>0 then
					if not xp_reselectflag then 
						response.write " disabled"
					else
						xp_qjlist = xp_qjlist&", "&idlist(yy-1)
						xp_qjsl = xp_qjsl&", "&sllist(yy-1)
					end if
				end if
				response.Write " checked><input type='hidden' name='inp_xp_qjsl' id='inp_xp_qjsl' value='"&sllist(yy-1)&"'>&nbsp;"
                response.Write rslistflag("yunyong")
                response.Write " - "&sllist(yy-1)
                'response.write "</div><div style='float:left; text-align:right; width:30%;'>"
                'response.write "&nbsp;说明&nbsp;<input type='text' name='inp_qjdesc' id='inp_qjdesc' size='4' value="""
                'if not isnull(rs("desc")) and rs("desc")<>"" then 
                '	if desclist(yy-1)<>"" then tmpdesc = desclist(yy-1)
                'end if
                'qjdesc=qjdesc & ", " & tmpdesc
                'response.write tmpdesc
                'response.write """>&nbsp;</div>"
                %></td>
          <%		if count22 mod 3 =0 then response.write "</tr><tr>"
			end if
		end if
		rslistflag.close()
	  next
	  if xp_qjlist<>"" then xp_qjlist = mid(xp_qjlist,3)
	  if xp_qjsl<>"" then xp_qjsl = mid(xp_qjsl,3)
	  %>
          <input type="hidden" name="xp_qjlist" id="xp_qjlist" value="<%=xp_qjlist%>">
          <input type="hidden" name="xp_qjsl" id="xp_qjsl" value="<%=xp_qjsl%>">
        </tr>
      </table>
    </div>
    <%end if%>
  </fieldset>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:5px; margin-bottom:5px">
  <tr>
    <td><input name="chk_isaddhq" type="checkbox" id="chk_isaddhq" value="yes" onClick="show_quick_addhq(this);">
隐藏后期&nbsp;&nbsp;总金额 <span id="lb_summoney">0</span> 元&nbsp; 现缴金额&nbsp;
<input name="txt_factmoney" type="text" id="txt_factmoney" size="6" onKeyUp="value=value.replace(/[^\d]/g,'')" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
元&nbsp;&nbsp;<a href="javascript:void(0);" onClick="javascript:openKeyPad(this)"><font color=blue><b>计算器</b></font></a><%if CheckOldMoneyControl() then%>
&nbsp;&nbsp;<span class="font">收款日期
<input name="savemoney_time" type="text" id="savemoney_time" value="<%=date()%>"  size="13" maxlength="10">
<a onClick="return showCalendar('savemoney_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG5" /></a></span>&nbsp;&nbsp;<%else
	response.write "<input name='savemoney_time' type='hidden' id='savemoney_time' value='"&date()&"'>"
end if%> <!--
<input name="chk_getallmoney" type="checkbox" id="chk_getallmoney" value="yes">
全额收款&nbsp;&nbsp;--> <input type="checkbox" name="wzsk" value="1">
刷卡收款&nbsp;( 优惠&nbsp;升级更换&nbsp; 朋友打折&nbsp;)</td>
    <td>&nbsp;</td>
  </tr>
</table>
<fieldset id="fs_hq">
<legend>后期消费</legend>
<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:5px">
  <%
	for i = 1 to 10
	%>
  <tr>
    <td width="540" height="22" style="meizz:expression(this.noWrap=true);">后期项目
      <select name="<%="top_type"&i%>" id="<%="top_type"&i%>" onChange="changelocation(this.options[this.selectedIndex].value,<%=i%>)">
          <option value="">请选择</option>
          <%set rs2=server.CreateObject("adodb.recordset")
			  rs2.open "select * from yunyong_type where ishidden=0 order by px asc",conn,1,1
			  while not rs2.eof 
			  response.Write "<option value='"&rs2("id")&"'>"&rs2("name")&"</option>"
			  rs2.movenext
			  wend
			  rs2.close
			  set rs2=nothing%>
        </select>
        <select name="<%="jixiang"&i%>" id="<%="jixiang"&i%>" style="width:150px" onChange="javascript:changePageShow(document.getElementById('pagevol<%=i%>'),this.options[this.selectedIndex].value,<%=i%>)">
          <option value="">请选择</option>
        </select>
      &nbsp; 总费用
      <input name="<%="money"&i%>" type="text" id="<%="money"&i%>" size="4" onKeyUp="value=value.replace(/[^\d]/g,'');getSumMoney()" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" onChange="getSumMoney()" onKeyDown="getSumMoney()">
      元 &nbsp;数量
      <input name="<%="sl"&i%>" type="text" id="<%="sl"&i%>" size="2" onKeyUp="value=value.replace(/[^\d]/g,'')   "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))"></td>
    <td id="<%="td_page"&i%>"><input name="<%="pagevol"&i%>" type="hidden" id="<%="pagevol"&i%>" size="2" onKeyUp="value=value.replace(/[^\d]/g,'')   "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))"></td>
    <td>&nbsp; 备注
      <input name="<%="beizhu"&i%>" type="text" id="<%="beizhu"&i%>" size="20"></td>
  </tr>
  <%next%>
</table>
</fieldset>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:5px; margin-bottom:5px">
  <tr>
    <td><input name="chk_xpxx" type="checkbox" id="chk_xpxx" value="yes" onClick="show_quick_xp(this);">
      隐藏选片报表</td>
    <td>&nbsp;</td>
  </tr>
</table>
<fieldset id="fs_xp">
<legend>选片信息</legend>
<iframe id="ewebeditor1" src="editor/ewebeditor.asp?id=content&style=s_blue" frameborder="0" scrolling="no" width="100%" height="250"></iframe>
<textarea name="content" id="content" style="display:none"></textarea>
</fieldset>
<div align="center">  <table width="97%" height="47"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="47" align="center">
      
	  <input name="tijiao" type="submit" id="tijiao" value="再次确定" onClick="return chk()">
  　	
    <input name="reset" type="button" id="reset" value="返回" onClick="javascript:history.go(-1)">
    <input name="id" type="hidden" id="id" value="<%=request("id")%>"></td>
    </tr>
    <tr>
      <td style="padding:5px"><p>给数码常用语<br>
        文件名    修眼睛  修胳膊  修脸型　修瘦　　修胖   牙齿修白等  大小眼　修肚子　　头发修饰　　不能修痣 <br>
        修景点杂物    脏的要去掉   重新设计   换背景  文字不要   　景点设计：　雪景　　草地　海边  瀑布　　淡化脑袋<br>
        定于月日取件！加急<br>
      </p>
        <p> 给摄影常用语<br>
          客人反应  取景不好 调色不太好  灯光问题   加把劲,　　让我们作更多的后期<br>
        </p>
        <p> 给门市常用语<br>
          件加急不了, 太多客人   定于月日给你    不能重设计了已输出了  　 与客人沟通能力用点技巧 <br>
        </p>
        <p>&nbsp;&nbsp;&nbsp;<span class="STYLE2">如果常用语不够,请写好文本传至公司技术部,我们将帮你们加上 </span></p></td>
    </tr>
  </table>
  </div>
</form>
<%rs.close
set rs=nothing
case else

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

<form action="xiadan_save.asp?action=save" method="post"  name="form1">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="xu_kuan">

    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" align="right" class="font">另选择下单时间：</td>
      <td class="font"><input name="times" type="text" id="times" value="<%=date%>" size="13" readonly>
        <a onClick="return showCalendar('times', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a></td>
      <td colspan="2" class="font">&nbsp;&nbsp;&nbsp;如果没另选择下单时间，默认时间为添加当天日期</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td width="108" height="30" align="right" class="font">摄影类型：</td>
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
 
  <%
  dim rs1
  set rs1=server.CreateObject("adodb.recordset")
	  rs1.open "select * from jixiang where ishidden=0 order by [type],px",conn,1,1
	  %>

  <%while not rs1.eof%>
  <option value="xiadan.asp?ids=<%=rs1("id")%>&id=<%=request("id")%>"><%=rs1("jixiang")%></option>
  <%rs1.movenext 
		wend 
		rs1.close
		set rs1=nothing%>
</select></td>
      <td width="94" align="right" class="font">套系金额：</td>
      <td width="266" class="font"><input name="money" type="text" id="money" size="13" value="<%if request("ids")<>"" then 
	  response.Write conn.execute("select money from jixiang where id="&request("ids")&"")(0)
	  end if%>" onKeyUp="value=value.replace(/[^\d]/g,'')   "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
      （元）&nbsp;&nbsp;&nbsp;
      <%if request("ids")<>"" then 
	  response.Write "原价："&conn.execute("select old_money from jixiang where id="&request("ids")&"")(0)
	  end if%></td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td width="108" height="31" align="right" class="font">摄影日期：</td>
      <td width="270" height="31" class="font"><input name="pz_time" type="text" maxlength="10" id="pz_time" size="13"/ >
        <a onClick="return showCalendar('pz_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="pz" type="text" size="3">
      点</td>
      <td align="right" class="font">拍照礼服：</td>
      <td class="font"><input name="pzlf_time" type="text" maxlength="10" id="pzlf_time" size="13" />
          <a onClick="return showCalendar('pzlf_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
          <input name="pzlf" id="pzlf" type="text" size="3">
  点 (可为空)</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" align="right" class="font">选片日期：</td>
      <td height="30" class="font"><a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#">
        <input name="kj_time" type="text" id="kj_time" size="13" >
      </a><a onClick="return showCalendar('kj_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
      <input name="kj" type="text" size="3">
      点</td>
      <td height="30" align="right" class="font">结婚礼服：</td>
      <td height="30" class="font"><a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#">
        <input name="jhlf_time" type="text" id="jhlf_time" size="13" >
      </a><a onClick="return showCalendar('jhlf_time', 'y-mm-dd');" href="#"><img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
      <input name="jhlf" type="text" id="jhlf" size="3">
      点 (可为空)</td>
    </tr>
	
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="30" align="right" class="font">取件日期：</td>
      <td height="30" class="font"><input name="qj_time" type="text" id="qj_time" size="13" >
          <a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#"> <img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
          <input name="qj" type="text" size="3">
          点</td>
      <td height="30" align="right" class="font">结婚化妆日期：</td>
      <td height="30" class="font"><a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#">
        <input name="hz_time" type="text" maxlength="10" id="hz_time" size="13" >
        </a><a onClick="return showCalendar('hz_time', 'y-mm-dd');" href="#"> <img src="Image/Button.gif" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
        <input name="hz" type="text" size="3">
        点</td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#ffffff">
      <td height="20" colspan="4" class="font">&nbsp;&nbsp;手动单号:
        <input name="danhao" type="text" id="danhao" size="8" onKeyUp="value=value.replace(/[^\d]/g,'')   "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
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
		  dim id1,sl1
		  id1=conn.execute("select yunyong from jixiang where id="&request("ids")&"")(0)
		  sl1=conn.execute("select sl from jixiang where id="&request("ids")&"")(0)
	  id=split(id1,", ")
	  sl=split(sl1,", ")
	 for ii=lbound(id) to ubound(id)
	 response.Write conn.execute("select yunyong from yunyong where id="&id(ii)&"")(0)&"["&sl(ii)&"]&nbsp;&nbsp;&nbsp;"
	 if ii=6 or ii=12 then
	 response.Write"<br>&nbsp;&nbsp;"
	 end if
	 next
	  
	  %>
</div></td>
    </tr>
	<%end if%>
</table>

  <%
  dim zz,rs3,rs4,tt
  set rs2=server.CreateObject("adodb.recordset")
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
	if not rs3.eof then
		
		%>
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
		response.write rs3("yunyong")
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
      <input name="queding" type="submit" id="确定" value="确定">
  　　　　　　　　
    <input name="reset" type="button" id="reset" value="返回" onClick="javascript:history.go(-1)">
    <input name="id" type="hidden" id="id" value="<%=request("id")%>">
      </div></td>
    </tr>
  </table>
  </div>
</form>
<%end select%>
<p>&nbsp;</p>
</body>
</html>

