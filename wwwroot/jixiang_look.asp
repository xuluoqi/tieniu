<!--#include file="zlsdk.asp"-->
<!--#include file="connstr.asp"-->
<!--include file="session.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/sms_class.asp"-->
<!--#include file="../inc/imgInfo.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="../Js/Calendar.js"></script>
<link href="../Css/imgzoom.css" rel="stylesheet" type="text/css">
<link href="../Css/calendar-blue.css" rel="stylesheet">
<link href="zxcss.css" rel="stylesheet" type="text/css">
<script src="../js/imgzoom.js" type="text/javascript"></script>
<script type="text/javascript">var IMGDIR = '/images';var attackevasive = '0';zoomstatus = parseInt(1);</script>
<STYLE>
<!--
A.ssmItems:link     {color:black;text-decoration:none;}
A.ssmItems:hover    {color:black;text-decoration:none;}
A.ssmItems:active   {color:black;text-decoration:none;}
A.ssmItems:visited  {color:black;text-decoration:none;}
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
.div_showprice{
	float:right;
	width:100px;
	cursor:pointer;
	color:#666666;
}
.initprice{
	display:none;
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
.pgtpbr_showl{z-index:2;position:absolute;top:0px;left:0px;width:30px;height:40px; background:url(../img/top_down.gif) no-repeat -5px -5px;cursor:pointer;}
.pgtpbr{
	z-index:1;
	position:absolute;
	top:0px;
	left:0px;
	width:30px;
	height:65px;
	background-color:#ffffff;
	background:url(../img/pgbar_bg.gif) no-repeat -8px -0px;
	cursor:pointer;
	border-bottom:solid 1px #91a9e1;
}
.pgtpbr1{
	z-index:1;
	position:absolute;
	top:0px;
	left:0px;
	width:30px;
	height:65px;
	background-color:#ffffff;
	background:url(../img/pgbar_bg1.gif) no-repeat -8px -0px;
	cursor:pointer;
	border-bottom:solid 1px #91a9e1;
}
-->
</STYLE>
<script language="javascript" src="../inc/ajax.js"></script>
<script language="javascript" src="../inc/func.js"></script>
<script language="javascript" src="../Js/jixiang_look.js"></SCRIPT>
<script language="javascript">
function onKeyDown(){  
	//event.ctrlKey && 
	if(event.keyCode==113){
		if(confirm("确定要关闭此窗口吗？")){window.close();}
	}
} 
document.onkeydown=onKeyDown;
</script>
<%
response.write "<script language='javascript'>"&vbcrlf
dim rstx,rstype,counts,flag,rows,rowcounts
dim tx_id,tx_yunyong,tx_sl,tx_type,tx_name,tx_money,tx_pagevol,tx_xc,rsyy,rsyytype
dim arr_yunyong,arr_sl,ct,st,modt
counts=0
st=0
set rstype = conn.execute("select * from companytype where ishidden=0")
set rstx = server.CreateObject("adodb.recordset")
response.write "mpmenu1=new mMenu('','javascript:void(0)','self','','','','');"&vbcrlf
do while not rstype.eof
	rstx.open "select * from jixiang where [type]="&rstype("id")&" and ishidden=0 order by px",conn,1,1
	counts=counts+1
	response.write "msub"&counts&"=new mMenuItem('"&rstype("companytype")&"','','self',false,'','1','','','','');"&vbcrlf
	if rstx.recordcount>0 then
		do while not rstx.eof
			response.write "msub"&counts&".addsubItem(new mMenuItem('"&rstx("jixiang")&"','jixiang_look.asp?id="&rstx("id")&"','self',false,'"&rstx("jixiang")&"',null,'','','',''));"&vbcrlf
			rstx.movenext
		loop
	else
		response.write "msub"&counts&".addsubItem(new mMenuItem('无','javascript:void(0);','self',false,'暂无套系',null,'','','',''));"&vbcrlf
	end if
	response.write "mpmenu1.addItem(msub"&counts&");"
	rstx.close
	rstype.movenext
loop
rstype.close
set rstype=nothing

'set rstype = conn.execute("select * from companytype where ishidden=0")
'set rstx = server.CreateObject("adodb.recordset")
'response.write "mpmenu2=new mMenu('','javascript:void(0)','self','','','','');"&vbcrlf
'do while not rstype.eof
'	rstx.open "select * from jixiang where [type]="&rstype("id")&" and ishidden=0 order by px",conn,1,1
'	counts=counts+1
'	response.write "msub"&counts&"=new mMenuItem('"&rstype("companytype")&"','','self',false,'','1','','','','');"&vbcrlf
'	if rstx.recordcount>0 then
'		do while not rstx.eof
'			tx_yunyong = replace(rstx("yunyong")," ","")
'			tx_sl = replace(rstx("sl")," ","")
'			tx_xc = ""
'			tx_id = ""
'			tx_name = ""
'			tx_type = ""
'			tx_money = ""
'			'arr_yunyong = split(tx_yunyong,",")
'			'arr_sl = split(tx_sl,",")
'			st=0
'			set rsyytype=conn.execute("select * from yunyong_type where ishidden=0 order by px asc")
'			do while not rsyytype.eof
'				set rsyy=conn.execute("select * from yunyong where type_id="&rsyytype("id")&" and ishidden=0 order by px asc")
'				do while not rsyy.eof
'					
'					st=st+1
'					if instr(","&tx_yunyong&",",","&rsyy("id")&",")>0 then
'						if st<10 then
'							st="00"&st
'						elseif st<100 then
'							st="0"&st
'						end if
'						tx_id = tx_id&","&st
'						tx_name = tx_name&","&rsyy("yunyong")
'						tx_type = tx_type&","&rsyy("type")
'						tx_xc = tx_xc&","&rsyy("isxc")
'						tx_money = tx_money&","&rsyy("money")
'					end if
'					rsyy.movenext
'				loop
'				modt=st mod 3
'				if modt>0 then st=st+3-modt
'				rsyy.close
'				set rsyy=nothing
'				rsyytype.movenext
'			loop
'			rsyytype.close
'			set rsyytype=nothing
'			
'			if tx_name<>"" then
'				tx_id = mid(tx_id,2)
'				tx_name = mid(tx_name,2)
'				tx_type = mid(tx_type,2)
'				tx_xc = mid(tx_xc,2)
'				tx_money = mid(tx_money,2)
'			end if
'			
'			response.write "msub"&counts&".addsubItem(new mMenuItem('"&rstx("jixiang")&"','javascript:ReloadContent(\'"&tx_yunyong&"\',\'"&tx_id&"\',\'"&tx_name&"\',\'"&tx_sl&"\',\'"&rstx("pagevol")&"\',\'"&tx_xc&"\',\'"&tx_type&"\',\'"&tx_money&"\')','self',false,'"&rstx("jixiang")&"',null,'','','',''));"&vbcrlf
'			rstx.movenext
'		loop
'	else
'		response.write "msub"&counts&".addsubItem(new mMenuItem('无','javascript:void(0);','self',false,'暂无套系',null,'','','',''));"&vbcrlf
'	end if
'	response.write "mpmenu2.addItem(msub"&counts&");"
'	rstx.close
'	rstype.movenext
'loop
'rstype.close
'set rstype=nothing
response.write "</script>"&vbcrlf

dim yunyong11,sl11,page11,pagevol,soption,arrtype,i,k,rs,rss,rs1,rs2,rs3,pz,kj,qj,cp,xg,hz,id
dim FSO,pic,gps,bFlag,DD,PWidth,PHeight,Pp,PXWidth,PXHeight,ImgSize,p1,sl,ii,imgpath
dim zz,a,namelist,typelist,moneylist,numlist,pagelist,xclist,tt,y,t3,x,sllist,counterlist,costlist
dim ver,pageInvisSetting,newOrderVerify,orderCostPoint,OrderCostControl,CameraInvis,DisableNewOrderDiscount

dim action,autosend,rsas
action=request.QueryString("action")

dim rs_setting
set rs_setting = conn.execute("select * from sysconfig")
if not (rs_setting.eof and rs_setting.bof) then
	ver = rs_setting("version")
	pageInvisSetting = rs_setting("pageInvisSetting")
	newOrderVerify = rs_setting("newOrderVerify")
	orderCostPoint = rs_setting("OrderCostPoint")
	OrderCostControl = rs_setting("OrderCostControl")
	CameraInvis = rs_setting("CameraInvis")
	DisableNewOrderDiscount = rs_setting("DisableNewOrderDiscount")
end if
rs_setting.close
set rs_setting = nothing

if session("level")=10 or (session("level")=1 and session("zhuguan")=1) then
	newOrderVerify=0
	OrderCostControl=0
end if

if action="add" then
	'添加客户
	dim lxpeple,telephone,address,address2,home_tel,telephone2,home_tel2,customerbeizhu,count11,savemoney
	dim sy_number,sys,beizhu11,danhao,temp,rsxd,kehu_id,userid2,userid3,xiangmu_id
	dim id3,ttt,desc
	
	lxpeple=request("lxpeple")
	telephone=request("telephone")
	address=request("address")
	address2=request("address2")
	home_tel=request("home_tel")
	telephone2=request("telephone2")
	home_tel2=request("home_tel2")
	customerbeizhu=htmlencode2(request("customerbeizhu"))
	if customerbeizhu="" or isnull(customerbeizhu) then customerbeizhu="&nbsp;"
	if trim(request("hqt_username"))<>"" then
		if conn.execute("select count(*) from kehu where hqt_username='"&trim(request("hqt_username"))&"'")(0)>0 then
			response.Write "<script>alert('婚庆通帐户名称已被人使用，请更换后重试！');history.back();</script>"
			Response.End
		end if
	end if
	
	''''''''''''套系检验''''''''''''''''''''''''''
	
	dim inp_oldid,inp_lxpeple,inp_lxpeple2,sqladd,editflag,addtimes
	 inp_oldid = request.form("inp_oldid")
	 inp_lxpeple = request.form("inp_lxpeple")
	 inp_lxpeple2 = request.form("inp_lxpeple2")
	 
	savemoney=request("savemoney")
	id=split(request("check"),", ")
	for i=lbound(id) to ubound(id)
		if not isnumeric(request("sl"&id(i)&"")) then
			response.Write "<script>alert('数量不能为空并且只能是数字，请检查！');history.back();</script>"
			Response.End
		end if
		sl11=sl11&request("sl"&id(i))&", "
		if request("p"&id(i))="" then
			page11=page11&"0, "
		else
			page11=page11&request("p"&id(i))&", "
		end if
		
		desc = desc & "|" & trim(request.form("desc"&id(i)))
	next
	if len(sl11)<=2 then
		response.Write "<script>alert('请至少选择一个套系内容，并填写数量！');history.back();</script>"
		Response.End
	else
		sl11=left(sl11,len(sl11)-2)
		page11=left(page11,len(page11)-2)
		desc=mid(desc,2)
	end if
	
	if telephone<>"" or home_tel<>"" or telephone2<>"" or home_tel2<>"" then 
		dim rscheck,sqlcheck
		if telephone<>"" then sqlcheck=sqlcheck&" or telephone='"&telephone&"' or home_tel='"&telephone&"' or telephone2='"&telephone&"' or home_tel2='"&telephone&"'"
		if home_tel<>"" then sqlcheck=sqlcheck&" or telephone='"&home_tel&"' or home_tel='"&home_tel&"' or telephone2='"&home_tel&"' or home_tel2='"&home_tel&"'"
		if telephone2<>"" then sqlcheck=sqlcheck&" or telephone='"&telephone2&"' or home_tel='"&telephone2&"' or telephone2='"&telephone2&"' or home_tel2='"&telephone2&"'"
		if home_tel2<>"" then sqlcheck=sqlcheck&" or telephone='"&home_tel2&"' or home_tel='"&home_tel2&"' or telephone2='"&home_tel2&"' or home_tel2='"&home_tel2&"'"
		sqlcheck=mid(sqlcheck,5)
		sqlcheck="select * from kehu where ("&sqlcheck&")"
		if inp_oldid<>"" and isnumeric(inp_oldid) then
			sqlcheck=sqlcheck&" and id<>"&inp_oldid
		end if
		set rscheck=server.createobject("adodb.recordset")
		rscheck.open sqlcheck,conn,1,1
		if not (rscheck.eof and rscheck.bof) then
			response.write "<script language='javascript'>alert('联系人手机或家庭电话号码重复,单击检测重复可查看具体信息.');history.back();</script>"
			response.end
		end if
		rscheck.close
		set rscheck= nothing
	end if
	
	dim ky_number,kynum
	if request("pz_time")<>"" then
		sys=conn.execute("select [CpMaxNum] from sysconfig")(0)
		if isnull(sys) then sys=0
		sy_number=conn.execute("select count(*) from shejixiadan where pz_time=#"&request("pz_time")&"#")(0)
		if sy_number>=sys and sys<>0 then
			response.Write "<script> alert('摄影当天已达到最高摄影人数,请另选择摄影日期！');history.back(); </script>"
			response.end  
		end if
	end if
	if request("kj_time")<>"" then
		kynum=conn.execute("select kyMaxNum from sysconfig")(0)
		if isnull(kynum) then kynum=0
		ky_number=conn.execute("select count(*) from shejixiadan where kj_time=#"&request("kj_time")&"#")(0)
		if ky_number>=kynum and kynum<>0 then
			response.Write "<script> alert('选片当天已达到最高选片人数,请另选择选片日期！');history.back(); </script>"
			response.end  
		end if
	end if
	
	beizhu11=request("beizhu")
	if ver="Customer" then 
		beizhu11=beizhu11&chr(10)&"套系金额："&request("money")&" 元"
		beizhu11=beizhu11&chr(10)&"选片后期："&request("txt_fujia")&" 元"
	end if
	  
	  if request("danhao")<>"" and isnumeric(request("danhao")) then
		danhao=conn.execute("select count(*) from shejixiadan where id="&request("danhao"))(0)
		if danhao>0 then
			response.Write "<script>alert('该单号已经存在，请检查单号是否错误！');history.back();</script>"
			Response.End
		end if
	  end if
	  
	 ''''''''''''套系检验''''''''''''''''''''''''''
	 
	 set rs=server.CreateObject("adodb.recordset")
	 
	 editflag=false
	 addtimes = cdate(Datevalue(cdate(trim(request("times"))))&" "&time())
	 
	 if inp_oldid<>"" and isnumeric(inp_oldid) then
	 	sqladd = "select * from kehu where id="&inp_oldid
		rs.open sqladd,conn,1,3
		if trim(inp_lxpeple)<>trim(lxpeple) and trim(inp_lxpeple)<>trim(request("lxpeple2")) and trim(inp_lxpeple2)<>trim(lxpeple) and trim(inp_lxpeple2)<>trim(request("lxpeple2")) then
			rs.addnew
		else
			editflag = true
		end if
	 else
	 	count11=conn.execute("select count(*) from kehu where [number]='"&request("number")&"' and [number]<>''")(0)
		if count11>0 then
			response.Write "<script>alert('该卡号跟已有卡号重复，请检查！');history.back();</script>"
			Response.End
		end if 
		sqladd = "select top 1 * from kehu"
		rs.open sqladd,conn,1,3
		rs.addnew
	 end if
	 rs("CustomerLostType")=request("CustomerLostType")
	
	  rs("number")=request("number")
	  rs("shopid")=request("shopid")
	  if request("js_id")<>"" and isnumeric(request("js_id")) then
	  	rs("js_id")=request("js_id")
	  end if
	  rs("group")=conn.execute("select [group] from yuangong where username='"&request("menshi")&"'")(0)
	 rs("lxpeple")=lxpeple
	if trim(telephone)<>"" and trim(telephone)<>"灵通号码前加区号" then
	 	rs("telephone")=telephone
	end if
	rs("qq")=trim(request("qq"))
	rs("qq2")=trim(request("qq2"))
	if trim(request("WeddingDay"))<>"" then
		rs("WeddingDay")=request("WeddingDay")
	end if
	if trim(request("PublishCardTime"))<>"" then
		rs("PublishCardTime")=request("PublishCardTime")
	end if
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
	 if request("chusheng2")<>"" and trim(request("chusheng2"))<>"格式如:8-8" then
		rs("chusheng2")=request("chusheng2")
	end if
	 rs("home_tel2")=request("home_tel2")
	 end if
	 rs("post")=request("post")
	 if trim(request("chusheng"))<>"" and trim(request("chusheng"))<>"格式如:8-8" then
		rs("chusheng")=trim(request("chusheng"))
	end if
	
	 rs("JhDateType")=request("JhDateType")
 	 rs("ShengriType")=request("ShengriType")
 	 rs("ShengriType2")=request("ShengriType2")
	 
	 if request("inp_oldid")="" then
	 	if trim(request("hqt_username"))<>"" and trim(request("hqt_password"))<>"" then
			rs("hqt_username")=trim(request("hqt_username"))
			rs("hqt_password")=MD5_16(trim(request("hqt_password")))
		end if
	else
		if trim(request("hqt_username"))="" and trim(request("hqt_password"))="" then
			rs("hqt_username")=null
			rs("hqt_password")=null
		else
			rs("hqt_username")=trim(request("hqt_username"))
			if trim(request("hqt_password"))<>"" then
				rs("hqt_password")=MD5_16(trim(request("hqt_password")))
			end if
		end if
	end if
	
	 rs("home_tel")=home_tel
	 rs("sex")=request("sex")
	 rs("shuoming")=customerbeizhu
	 rs("c_pic")=request("c_pic")         '新增客户上传图片
	 rs("userid")=request("menshi")
	 rs("userid2")=request("menshi2")
	 rs("userid3")=request("menshi3")
	 if not editflag then rs("times")=addtimes
	 rs("pianhao")=request("check2")
	 rs("islost")=0
	 rs.update
	 
	 if not editflag then
		 temp = rs.bookmark
		 rs.bookmark = temp
	 end if
	 kehu_id=rs("ID")                '客户ID
	 rs.close()
	 
	 if kehu_id="" or not isnumeric(kehu_id) then
	 	response.Write "<script>alert('客户资料添加失败，请重新操作！');history.back();</script>"
	  	Response.End
	 end if
	 
	 '客户添加完毕
	 
	'预约
	 set rsxd=server.CreateObject("adodb.recordset")
	rsxd.open "select * from shejixiadan ",conn,1,3
	rsxd.addnew 
	rsxd("danhao")=request("danhao")
	rsxd("kehu_id")=kehu_id
	rsxd("sl")=sl11
	rsxd("pagevol")=page11
	rsxd("yunyong")=request("check")
	rsxd("desc")=desc
	rsxd("jixiang")=request("jixiang")
	if request("hz_time")<>"" and request("hz")<>"" then
	rsxd("hz_time")=request("hz_time")
	rsxd("hz")=request("hz")
	end if
	rsxd("beizhu")=htmlencode2(beizhu11)
	if ver="Customer" then 
		rsxd("jixiang_money")=0
	else
		rsxd("jixiang_money")=request("money")
	end if
	if request("hz_time")<>"" and request("hz")<>"" then
		rsxd("hz_time")=request("hz_time")
		rsxd("hz")=request("hz")
	end if
	if request("pz_time")<>"" and request("pz")<>"" then
		rsxd("pz_time")=request("pz_time")
		rsxd("pz")=request("pz")
	end if 
	if request("hhz_time")<>"" and request("hhz")<>"" then
		rsxd("hhz_time")=request("hhz_time")
		rsxd("hhz")=request("hhz")
	end if 
	if request("pz_time2")<>"" and request("pz2")<>"" then
		rsxd("pz_time2")=request("pz_time2")
		rsxd("pz2")=request("pz2")
	end if 
	if request("pzlf_time")<>"" and request("pz")<>"" then
		rsxd("pzlf_time")=request("pzlf_time")
		rsxd("pzlf")=request("pzlf")
	end if
	if request("jhlf_time")<>"" and request("jhlf")<>"" then
		rsxd("jhlf_time")=request("jhlf_time")
		rsxd("jhlf")=request("jhlf")
	end if
	if request("kj_time")<>"" and request("kj")<>"" then
		rsxd("kj_time")=request("kj_time")
		rsxd("kj")=request("kj")
	end if
	
	if request("qj_time")<>"" and request("qj")<>"" then
		rsxd("qj_time")=request("qj_time")
		rsxd("qj")=request("qj")
	end if
	if request("xg_time")<>"" and request("xg")<>"" then
		rsxd("xg_time")=request("xg_time")
		rsxd("xg")=request("xg")
	end if
	
	
	rsxd("group")=conn.execute("select [group] from yuangong where username='"&request("menshi")&"'")(0)
	rsxd("userid")=request("menshi")
	userid2=request("menshi2")
	if userid2<>"" then 
	rsxd("userid2")=userid2
	else
	rsxd("userid2")=""
	end if
	userid3=request("menshi3")
	if userid3<>"" then 
	rsxd("userid3")=userid3
	else
	rsxd("userid3")=""
	end if
	rsxd("kj_userid")=request("menshi")
	if request("times")="" then
	rsxd("times")=date()
	else
	rsxd("times")=addtimes
	
	end if
	rsxd("stated")=request("stated")
	rsxd("sl2")=request("sl2")
	
	'快速安排员工
	if request("chk_quick_arrangements") = "yes" then
		if request("hz_name")<>"" then
			rsxd("hz_name") = conn.execute("select peplename from yuangong where username='"&request("hz_name")&"'")(0)
			rsxd("lc_hz") = addtimes
		end if
		if request("cp_name")<>"" then
			rsxd("cp_name") = conn.execute("select peplename from yuangong where username='"&request("cp_name")&"'")(0)
			rsxd("lc_cp") = addtimes
			if request("cp_wedvol")<>"" and isnumeric(request("cp_wedvol")) then
				rsxd("cp_wedvol") = request("cp_wedvol")
			end if
			rsxd("cp_memo") = request("cp_memo")
		end if
		if request("cp_name2")<>"" then
			rsxd("cp_name2") = conn.execute("select peplename from yuangong where username='"&request("cp_name2")&"'")(0)
			rsxd("lc_cp") = addtimes
			if request("cp_wedvol2")<>"" and isnumeric(request("cp_wedvol2")) then
				rsxd("cp_wedvol2") = request("cp_wedvol2")
			end if
			rsxd("cp_memo2") = request("cp_memo2")
		end if
		if request("ts_name")<>"" then
			rsxd("xp_name") = conn.execute("select peplename from yuangong where username='"&request("ts_name")&"'")(0)
			rsxd("lc_xp") = addtimes
			if request("tsVolume")<>"" and isnumeric(request("tsVolume")) then
				rsxd("tsVolume") = Cint(request("tsVolume"))
			end if
			if request("cpVolume")<>"" and isnumeric(request("cpVolume")) then
				rsxd("cpVolume") = Cint(request("cpVolume"))
			end if
		end if
		'if request("xp_name")<>"" then
		'	rsxd("kj_userid")=request("xp_name")
		'	rsxd("ky_name") = conn.execute("select peplename from yuangong where username='"&request("xp_name")&"'")(0)
		'	rsxd("lc_ky") = addtimes
		'end if
		if request("hz_name2")<>"" then
			rsxd("hz_name2") = request("hz_name2")
		end if
		if request("cpzl_name")<>"" then
			rsxd("cpzl_name") = request("cpzl_name")
		end if
		'if request("sj_name")<>"" then
		'	rsxd("sj_name") = conn.execute("select peplename from yuangong where username='"&request("sj_name")&"'")(0)
		'	rsxd("lc_sj") = addtimes
		'end if
		if request("qujian") = "yes" then
			rsxd("wc_name") = conn.execute("select peplename from yuangong where username='"&request("menshi")&"'")(0)
			rsxd("lc_wc") = addtimes
		end if
	end if
	
	rsxd("jhz_style") = request("jhz_style")
	if session("adminid")<>"" and isnumeric(session("adminid")) then
		rsxd("createuserid") = session("adminid")
	elseif session("zg_adminid")<>"" and isnumeric(session("zg_adminid")) then
		rsxd("createuserid") = session("zg_adminid")
	end if	
	rsxd.update
	temp = rsxd.bookmark
	rsxd.bookmark = temp
	xiangmu_id=rsxd("ID")                'xiangmu_id
	rsxd.close
	set rsxd=nothing
	
	'快速安排员工
	if request("chk_quick_arrangements") = "yes" then
		if request("hz_name")<>"" then
			conn.execute("insert into xiadan (userid,beizhu,xiangmu_id,type,times) values ('"&request("hz_name")&"','&nbsp;',"&xiangmu_id&",5,#"&addtimes&"#)")
		end if
		if request("cp_name")<>"" then
			conn.execute("insert into xiadan (userid,beizhu,xiangmu_id,type,times) values ('"&request("cp_name")&"','&nbsp;',"&xiangmu_id&",4,#"&addtimes&"#)")
		end if
		if request("cp_name2")<>"" then
			conn.execute("insert into xiadan (userid,beizhu,xiangmu_id,type,times) values ('"&request("cp_name2")&"','&nbsp;',"&xiangmu_id&",4,#"&addtimes&"#)")
		end if
		if request("cp_name3")<>"" then
			conn.execute("insert into xiadan (userid,beizhu,xiangmu_id,type,times) values ('"&request("cp_name3")&"','&nbsp;',"&xiangmu_id&",4,#"&addtimes&"#)")
		end if
		if request("sj_name")<>"" then
			conn.execute("insert into xiadan (userid,beizhu,xiangmu_id,type,times) values ('"&request("sj_name")&"','&nbsp;',"&xiangmu_id&",2,#"&addtimes&"#)")
		end if
	end if
	
	'接单确认取件不扣库存
	if request("qujian") <> "yes" then
		id3=split(request("check"),", ")
		ttt=split(sl11,", ")
		
		for ii=lbound(id3) to ubound(id3)
			set rsyy = conn.execute("select [type] from yunyong where id="&id3(ii))
			if not rsyy.eof then
				if rsyy(0)=1 then
					conn.execute("update yunyong set sl=sl-"&ttt(ii)&" where id="&id3(ii)&"")
					conn.execute("insert into cuenchu (xiangmu_id,sp_id,sl,type,type2,type3,beizhu,times) values ("&xiangmu_id&","&id3(ii)&","&ttt(ii)&",2,1,1,'"&htmlencode2(request("beizhu"))&"',#"&addtimes&"#)")
				end if
			end if
		next
	end if
	
	'预收套系金额
	if savemoney>0 then
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from save_money",conn,1,3
		rs.addnew
		rs("userid")=request("menshi")
		rs("group")=conn.execute("select [group] from yuangong where username='"&request("menshi")&"'")(0)
		rs("xiangmu_id")=xiangmu_id
		rs("money")=savemoney
		rs("type")=1
		if request("wzsk")="yes" then
			rs("wzsk")=1
		else
			rs("wzsk")=0
		end if
		if request("times")="" then
			rs("times")=now()
		else
			rs("times")=addtimes
		end if
		rs("beizhu")="定金付款"
		rs.update
		temp = rs.bookmark
		rs.bookmark = temp
		dim save_id:save_id=rs("ID")                'save_id
		rs.close
		set rs=nothing
	end if
	Call FinalMoneySum(xiangmu_id,True)
	
	'保存预设时间记录
	'dim dict_time
'	set dict_time=Server.CreateObject("Scripting.Dictionary")
'	dict_time("hz")=request("hz_time")
'	dict_time("pz")=request("pz_time")
'	dict_time("pz2")=request("pz_time2")
'	dict_time("kj")=request("kj_time")
'	dict_time("xg")=request("xg_time")
'	dict_time("hhz")=request("hhz_time")
'	dict_time("pzlf")=request("pzlf_time")
'	dict_time("jhlf")=request("jhlf_time")
'	dict_time("qj")=request("qj_time")
'	Call EditedTimeSaveToReport(xiangmu_id,0,"hz",null,request("hz_time"))
	
	
	'客怨
	if request.form("kyflag")="yes" then
		dim rsvote
		set rsvote= server.createobject("adodb.recordset")
		sqlvote = "select top 1 * from vote"
		rsvote.open sqlvote,conn,1,3
		rsvote.addnew
		rsvote("xiangmu_id") = xiangmu_id
		rsvote("kehu_id") = kehu_id
		if request.form("kyflag")="yes" then
			rsvote("kyflag") = 1
		else
			rsvote("kyflag") = 0
		end if
		rsvote("memo") = request.form("kymemo")
		rsvote.update
		set rsvote = nothing
	end if
	
	'接单自动发送短信
	if request.form("chk_autosend")="yes" then
		dim un
		un = conn.execute("select peplename from yuangong where username='"&request("menshi")&"'")(0)
		Call SMSAutoPost("new",0,kehu_id,un)
	end if
	if request("js_id")<>"" and isnumeric(request("js_id")) then
	  	Call SMSAutoPost("js",0,kehu_id,un)
	end if
	
'	if err.number>0 then
'		conn.execute("delete from save_money where id="&save_id)
'		response.Write "<script language=javascript>alert('套系预约下单失败.\n"&err.description&"'.);history.back();</ script>"
'	else
		response.Write "<script language=javascript>"
		response.write "alert('套系预约成功，订单单号为"& xiangmu_id &"。');"
		response.write "location.href='lc_baobiao.asp?id="& xiangmu_id &"';"
		response.write "</script>"
'	end if
	response.end
end if

dim rstx_info,sqltx,jhz_style,jixiang_name,jixiang_money
id = request("id")
if id="" or not isnumeric(id) then
	sqltx="select top 1 * from jixiang"
else
	sqltx="select * from jixiang where id="&id
end if
set rstx_info = conn.execute(sqltx)
if not (rstx_info.eof and rstx_info.bof) Then
	id = rstx_info("id")
	jixiang_name = rstx_info("jixiang")
	jixiang_money = rstx_info("money")
	yunyong11 = rstx_info("yunyong")
	sl11 = rstx_info("sl")
	pagevol = rstx_info("pagevol")
	jhz_style = rstx_info("jhz_style")
	
	If jixiang_money >= 0 Then 
		response.write "<script language='javascript'>"&vbcrlf
		response.write "orderOldMoney="&jixiang_money&";"&vbcrlf
		If session("level")=10 Or session("level")=1 Or session("level")=7 Then 
			response.write "DLCheckUserLevel('"& session("level") &"','"& session("zhuguan") &"');"&vbcrlf
		Else
			response.write "DLCheckUserLevel('','');"&vbcrlf
		End If 
		response.write "DiscountListInit();"&vbcrlf

		If DisableNewOrderDiscount = 1 Then 
			Dim disc
			For disc = 1 To 4
				If rstx_info("NewOrderDiscount"&disc) > 0 Then 
					response.write "DiscountListAdd(new Array("& disc &","& rstx_info("NewOrderDiscount"&disc) &","& rstx_info("NewOrderDiscount"&disc)*jixiang_money/100 &"));"&vbcrlf
				End If 
			Next 
		End If 
		response.write "</script>"
	End If 
else
	response.write "<script language=javascript>alert('参数错误,本窗口将自动关闭.');window.close();</script>"
	Response.End
end if
%>
<title><%=jixiang_name%> - 婚纱摄影系统</title>
<SCRIPT LANGUAGE="VBScript"> 
Sub ExecShell(FilePath)
	on error resume next
	Dim WshShell
	Dim q
	Dim sCmd
	Dim Ret
	Set WshShell = CreateObject("WScript.Shell")
	q = Chr(34)
	sCmd=q & FilePath & q
	Ret = WshShell.Run (sCmd,0,true)
	error.clear()
End Sub
</script>
<script language="javascript">
var arrItem;
for(var i=0;i<arrDiscountList.length;i++){
	arrItem = arrDiscountList[i];
}
//alert(text);
</script>
</head>
<body onLoad="page_load()" onResize="setResize()">
<div id="append_parent"></div><div id="ajaxwaitid"></div>
<div class="pgtpbr_showl" id="top_down" onClick="pageTopBar_show();">&nbsp;</div>
<div class="pgtpbr" id="topbar"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  	<td width="85" height="63" onClick="javascript:document.all('div_content').scrollTop=0;"></td>
    <td width="75" onClick="javascript:document.all('div_content').scrollTop=getAbsTop(document.all('tb_order_info'))-70;"></td>
    <td width="75" onClick="javascript:window.location.href='#bottom'"></td>
    <td width="70" onClick="javascript:openwindow('../mxb.asp',400,300);"></td>
    <td width="80" valign="bottom"><script language="javascript">mwritetodocument();</script></td>
    <td width="80" onClick="<%
	pic="../upload/"&rstx_info("pic")
	'check file
	
	if pic="" or isnull(pic) then
		response.write "javascript:alert('暂时无图片预览.');"
	else
		set FSO=server.createobject("scripting.filesystemobject")
		imgpath=server.mappath(pic)
		if FSO.FileExists(imgpath) then
			response.Write "javascript:zoom(this, '"&pic&"');"
		else
			response.write "javascript:alert('暂时无图片预览.');"
		end if
		set FSO = nothing
	end if
	%>"></td>
    <td width="70" onClick="javascript:if(confirm('确定要关闭此窗口吗？')){window.close();}"></td>
    <td><div style="border:dashed 1px #cccccc; width:400px; margin:4px 2px 2px 2px; padding:3px; height:52px"><%
		response.write "<strong>&nbsp;"&rstx_info("jixiang")&"</strong>"
		response.write "<font color='#999999'>&nbsp;&nbsp;&nbsp;&nbsp;价格："&rstx_info("money")&" 元<br>"
		response.write "&nbsp;点击图片放大即可调出同类所有图片</font>"
	%></div></td>
  </tr>
</table></div>
<div id="div_content" style="height:100%; width:100%; overflow-y:auto">
<form action="jixiang_look.asp?action=add" method="post" name="form1">
<div id="div_blank" style="height:12px;"></div>
<script language="javascript">setResize()</script>
<table width="979" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="51" background="../img/order_01.gif"><table id="tb_1" width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr id="tr_addtitle">
        <td width="80%" height="30" align="center">[图片即为本套系内容 &nbsp;&nbsp;点击图片，可调出同类图片]
          <input type="hidden" name="jixiang" value="<%=id%>" /></td>
        <td width="14%" align="right" id="td_addcus"><a href="#" onClick="div_show()"><b><font color="red">·添加客户资料并预约</font></b></a></td>
        <td width="6%" align="center" id="td_addcus"><a href="javascript:window.close()"><b>·关闭</b></a></td>
      </tr>
    </table>
	<table id="tb_2" width="100%"  border="0" align="center" cellpadding="0" cellspacing="0" style="display:none">
	  <tr>
		<td width="80%" height="30"><div align="center"><strong><%=rstx_info("jixiang")%></strong>&nbsp;包含以下&nbsp;&nbsp;[请点击产品查图片]</div></td>
		<td width="14%" align="right"><a href="#" onClick="div_hidden()"><strong><font color="red">·取消此操作</font></strong></a></td>
	    <td width="6%" align="center">·<a href="javascript:window.close()"><b>关闭</b></a></td>
	  </tr>
	</table>
      </td>
  </tr>
</table>
<div id="div_customer" style=""><!-- style="display:none"-->
<div style="width:975px; background-color:#FFFFFF; background-position:right top;  background-repeat:no-repeat">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
    <tr valign="middle"><td colspan="33" height="5"></td></tr>
    <tr valign="middle">
      <td height="20" colspan="3" class="font">&nbsp;<%Call ShowMultipleShopSelect(1,Session("UserShopID"),true)%>
       &nbsp;门市1
        <select name="menshi" id="menshi">
        	<option value="">请选择</option>
              <%set rs=conn.execute("select username,peplename from yuangong where [level]=1 and isdisabled=0")
				while not rs.eof %>
              <option value="<%=rs("username")%>" <%if session("level")=1 then 
			  	if rs("username")=session("userid") then response.write "selected"
			  end if %>><%=rs("peplename")%></option>
              <%rs.movenext
			wend 
			rs.close
			set rs=nothing
			%>
            </select>
门市2
         <select name="menshi2" id="menshi2">
  <option value="">请选择</option>
  <%set rs=conn.execute("select username,peplename from yuangong where [level]=1 and isdisabled=0")
	while not rs.eof %>
  <option value="<%=rs("username")%>"><%=rs("peplename")%></option>
  <%rs.movenext
	wend 
	rs.close
	set rs=nothing%>
</select>
其他员工
<select name="menshi3" id="menshi3">
  <option value="">请选择</option>
  <%
  set rs_ygtype = conn.execute("select * from worktype where [level]<>1 order by id")
  do while not rs_ygtype.eof
  	%>
    <OPTGROUP LABEL="<%=GetDutyName(rs_ygtype("level"))%>">
    <%
		set rs=conn.execute("select * from yuangong where [level]="&rs_ygtype("level")&" and isdisabled=0")
		do while not rs.eof
	%>
    		<option value="<%=rs("username")%>"><%=rs("peplename")%></option>
	<%
		rs.movenext
	loop
	rs.close
	set rs=nothing
	%>
	</OPTGROUP>
	<%
  	rs_ygtype.movenext
  loop
  rs_ygtype.close
  set rs_ygtype=nothing
  %>
</select>
拍照类型
<select name="CustomerLostType" id="CustomerLostType">
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
</select></td>
      <td width="32%" rowspan="8" valign="top" class="font" <%if CameraInvis=0 then response.write "style='background-image:url(../img/order_03.gif); background-position:top right;background-repeat:no-repeat'"%>><%if CameraInvis=1 then%><iframe src="../pdt/index.asp" width="300" height="277" frameborder="0" scrolling="no"></iframe><%end if%>&nbsp;</td>
    </tr>
    <tr align="left" valign="middle">
      <td height="20" colspan="3" class="font">&nbsp;VIP卡号:
        <input name="number" type="text" id="number" size="10" />
  &nbsp; 发卡日期:
        
        <input name="PublishCardTime" type="text" id="PublishCardTime" size="14" />
        <a onClick="return showCalendar('PublishCardTime', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG" /></a>     &nbsp;&nbsp;结婚纪念:
        <select name="JhDateType" id="JhDateType">
          <option value="0" selected>农历</option>
          <option value="1">公历</option>
        </select>
        <input name="WeddingDay" type="text" id="WeddingDay" size="14" />
        <a onClick="return showCalendar('WeddingDay', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG" /></a><br>
        &nbsp;介绍人:&nbsp;&nbsp;
        <input name="js_name" type="text" id="js_name" size="10" onClick="javascript:openkhwidnow();" readonly />
<input type="button" name="button" id="button" value="清空" style="width:30px; background-color:eee" onClick="javascript:$E('js_name').value='';$E('js_id').value='';">
<input type="hidden" name="js_id" id="js_id">&nbsp;&nbsp;图片   
     <input name="c_pic" type="text" id="c_pic" size="25" />
        <a href="#" onClick="window.open('upfile.asp?formname=form1&editname=c_pic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">上传图片</a><input name="addtimes" type="hidden" id="addtimes" value="<%=now()%>" /></td>
      </tr>
    <tr align="left" valign="middle">
      <td height="20" colspan="3" class="font">&nbsp;<input name="chk_quick_arrangements" type="checkbox" id="chk_quick_arrangements" value="yes" onClick="show_quick_arrangements(this);">
        快速安排员工<font color=#AAAAAA>（如已出件请选择确认取件，如果是客服软件会自动确认取件）</font>
         &nbsp;&nbsp;&nbsp; <input name="qujian" type="checkbox" id="qujian" value="yes">
确认取件&nbsp;&nbsp;&nbsp;&nbsp; <input type="checkbox" name="chk_autosend" id="chk_autosend" value="yes"<%
autosend = GetAutoPostFlag("new")
select case autosend
	case 1
		response.write " checked"
	case -1
		response.write " disabled title='未配置接单短信设置'"
end select
%>>
信息</td>
      </tr>
    <tr align="left" valign="middle" id="tr_quick_arrangements" style="display:none">
      <td height="20" colspan="3" class="font">
        &nbsp;拍照化妆
          <select name="hz_name" id="hz_name">
          <option value="">请选择...</option>
          <%
			  set rss = server.CreateObject("adodb.recordset")
			  rss.open "select * from yuangong where level=5 and isdisabled=0",conn,1,1
			  do while not rss.eof
			  %>
          <option value="<%=rss("username")%>"><%=rss("peplename")%></option>
          <%
			  rss.movenext
			  loop
			  rss.close
			  %>
        </select>
       &nbsp;&nbsp;<%=GetDutyName(14)%>
<select name="hz_name2" id="hz_name2">
  <option value="">请选择...</option>
  <%
			  set rss = server.CreateObject("adodb.recordset")
			  rss.open "select * from yuangong where level=14 and isdisabled=0",conn,1,1
			  do while not rss.eof
			  %>
  <option value="<%=rss("peplename")%>"><%=rss("peplename")%></option>
  <%
			  rss.movenext
			  loop
			  rss.close
			  %>
</select>
&nbsp;&nbsp;摄影助理
<select name="cpzl_name" id="cpzl_name">
  <option value="">请选择...</option>
  <%
			  set rss = server.CreateObject("adodb.recordset")
			  rss.open "select * from yuangong where level=12 and isdisabled=0",conn,1,1
			  do while not rss.eof
			  %>
  <option value="<%=rss("peplename")%>"><%=rss("peplename")%></option>
  <%
			  rss.movenext
			  loop
			  rss.close
			  %>
</select>
<br>
       &nbsp;摄影师 1
         <select name="cp_name" id="cp_name">
          <option value="">请选择...</option>
          <%
			  set rss = server.CreateObject("adodb.recordset")
			  rss.open "select * from yuangong where level=4 and isdisabled=0",conn,1,1
			  do while not rss.eof
			  %>
          <option value="<%=rss("username")%>"><%=rss("peplename")%></option>
          <%
			  rss.movenext
			  loop
			  rss.close
			  %>
        </select>
        照片
        <input name="cp_wedvol" type="text" id="cp_wedvol" value="0" size="3" maxlength="3">
          说明
          <input name="cp_memo" type="text" id="cp_memo" size="50">
          <br>
        &nbsp;摄影师 2
         <select name="cp_name2" id="cp_name2">
          <option value="">请选择...</option>
          <%
			  set rss = server.CreateObject("adodb.recordset")
			  rss.open "select * from yuangong where level=4 and isdisabled=0",conn,1,1
			  do while not rss.eof
			  %>
          <option value="<%=rss("username")%>"><%=rss("peplename")%></option>
          <%
			  rss.movenext
			  loop
			  rss.close
			  %>
        </select>
        照片
        <input name="cp_wedvol2" type="text" id="cp_wedvol2" value="0" size="3" maxlength="3">
        说明
        <input name="cp_memo2" type="text" id="cp_memo2" size="50">
<br>
&nbsp;调色　　
<select name="ts_name" id="ts_name">
  <option value="">请选择...</option>
  <%
  dim rs_ygtype,rs_yginfo
  set rs_ygtype = conn.execute("select * from worktype where [level]=2 or [level]=4 or [level]=12 order by id")
  do while not rs_ygtype.eof
	%>
	<OPTGROUP LABEL="<%=GetDutyName(rs_ygtype("level"))%>">
	<%
		set rs_yginfo=conn.execute("select * from yuangong where [level]="&rs_ygtype("level")&" and isdisabled=0")
		do while not rs_yginfo.eof
	%>
			<option value="<%=rs_yginfo("username")%>"><%=rs_yginfo("peplename")%></option>
	<%
		rs_yginfo.movenext
	loop
	rs_yginfo.close
	set rs_yginfo=nothing
	%>
	</OPTGROUP>
	<%
	rs_ygtype.movenext
  loop
  rs_ygtype.close
  set rs_ygtype=nothing
  %>
</select>
修片张数
<input name="tsVolume" type="text" id="tsVolume" size="10">
&nbsp;摄影张数
<input name="cpVolume" type="text" id="cpVolume" size="10">
<br>
&nbsp;是否客怨
<input name="kyflag" type="checkbox" id="kyflag" value="yes">
&nbsp;&nbsp; 客怨说明
<input name="kymemo" type="text" id="kymemo" size="50"></td>
      </tr>
    <tr align="left" valign="middle">
      <td colspan="3" class="font"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
        <tr align="left" valign="middle">
          <td height="35" colspan="3" class="font"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="45%"><img src="../img/order_05.gif" width="440" height="28" /></td>
                <td width="55%" style="padding-left:110px"><input type="button" name="button2" id="button2" value="查找已有客户" style="background-color:eeeeee; width:100px; height:22px" onClick="javascript:op();">
                    <input name="inp_oldid" type="hidden" id="inp_oldid" />
                    <input name="inp_lxpeple" type="hidden" id="inp_lxpeple" />
                    <input name="inp_lxpeple2" type="hidden" id="inp_lxpeple2" /></td>
              </tr>
          </table></td>
        </tr>
        <tr align="left" valign="middle">
          <td width="33%" height="22" class="font">&nbsp;<%=GetAppellation(1)%>:
            <input name="lxpeple" type="text" id="lxpeple" size="20" />          </td>
          <td width="32%" height="22" class="font">&nbsp;客人性别:
            <input name="sex" type="radio" value="男" checked="checked" />
            男
            <input type="radio" name="sex" value="女" />
            女</td>
          <td width="35%" height="22" class="font">&nbsp;出生年月:
            <select name="ShengriType" id="ShengriType">
                <option value="0" selected>农历</option>
                <option value="1">公历</option>
              </select>
              <input name="chusheng" type="text" id="chusheng" size="12" class="inp1" onFocus="if (this.value == this.defaultValue) this.value='';" onBlur="if (this.value==''){this.value=this.defaultValue;}else{CheckIsShortDate(this,'请输入正确的生日日期格式,如:8-8或<%=year(date())-25%>-8-8.\t')}" value="格式如:8-8"></td>
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
                  <input name="address" type="text" id="address" size="40" />
          </div></td>
          <td height="22" class="font">&nbsp;邮政编码:
            <input name="post" type="text" id="post5" size="20" /></td>
        </tr>
      </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr align="left" valign="middle">
            <td height="35" colspan="3" class="font"><img src="../img/order_04.gif" width="440" height="28" /></td>
          </tr>
          <tr align="left" valign="middle">
            <td width="33%" height="20" class="font">&nbsp;<%=GetAppellation(2)%>:
              <input name="lxpeple2" type="text" id="lxpeple22" size="20" />
                <div align="right"></div></td>
            <td width="32%" height="20" class="font">&nbsp;客人性别:
              <input name="sex2" type="radio" value="男" />
              男
              <input name="sex2" type="radio" value="女" checked="checked" />
              女 </td>
            <td width="35%" height="20" class="font">&nbsp;出生年月:
              <select name="ShengriType2" id="ShengriType2">
                  <option value="0" selected>农历</option>
                  <option value="1">公历</option>
                </select>
                <input name="chusheng2" type="text" id="chusheng2" size="12" class="inp1" onFocus="if (this.value == this.defaultValue) this.value='';" onBlur="if (this.value==''){this.value=this.defaultValue;}else{CheckIsShortDate(this,'请输入正确的生日日期格式,如:8-8或<%=year(date())-25%>-8-8.\t')}" value="格式如:8-8"></td>
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
                    <input name="address2" type="text" id="address2" size="40" />
                </div>
              <div align="left"></div>
            </div></td>
            <td height="20" class="font">&nbsp;邮政编码:
              <input name="post2" type="text" id="post2" size="20" /></td>
          </tr>
          <tr align="left" valign="top">
            <td height="19" colspan="3"><table width="562" height="39" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="60" height="39" valign="top" class="font">&nbsp;备注:</td>
                  <td valign="top" class="font"><textarea name="customerbeizhu" cols="67" rows="2" id="customerbeizhu"></textarea></td>
                </tr>
            </table></td>
          </tr>
          <tr align="left" valign="top">
            <td height="19" colspan="3"><img src="../img/order_06.gif" width="440" height="28"></td>
          </tr>
          <tr>
            <td height="30">&nbsp;用户名：
              <input name="hqt_username" type="text" id="hqt_username" size="15"></td>
            <td height="30">&nbsp;密码：
              <input name="hqt_password" type="password" id="hqt_password" size="20"></td>
            <td height="30"><div id="hqt_msg"></div></td>
          </tr>
        </table></td>
      </tr>
	</table>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
    <tr align="left" valign="middle" bgcolor="#FFFFFF">
      <td height="40" colspan="3" class="font" onClick="javascript:showPianhaoDetails()" style="cursor:pointer" title="单击显示/隐藏个人沟通调查"><img src="../img/order_02.gif" width="702" height="32" /></td>
    </tr>
    <tr align="left" valign="top" id="tr_pianhao_details" style="display:none">
      <td height="31" colspan="3"><%set rs3=server.CreateObject("adodb.recordset")
	rs3.open "select * from pianhao where ishidden=0 order by px,id",conn,1,1
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
	rs.open "select * from pianhao_list where title_id="&rs3("id")&" and ishidden=0 order by px,id",conn,1,1
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
</div>
<br>
<table id="tb_order_info" width="975" height="283" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="xu_kuan">
  <tr align="left" valign="middle" bgcolor="#FFFFFF">
    <td width="150" height="30" align="right" class="font">下单时间：</td>
    <td class="font"><%if CheckOldMoneyControl() then%><input name="times" type="text" id="times" value="<%=date%>" size="13" onChange="CheckDateInfo('times','sp_times')" onBlur="CheckDateInfo('times','sp_times')" />
      <a onClick="return showCalendar('times', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a><%else
	  response.write "<input name='times' type='text' id='times' value='"&now()&"' size='20' readonly>"
	end if%>
      &nbsp;&nbsp;&nbsp;<a href="javascript:void(0)" onClick="javascript:openwindow('huensha_show.asp',800,600);" ><font color=red><b>礼服列表</b></font></a>&nbsp;&nbsp;<a href="javascript:void(0);" onClick="javascript:openKeyPad(this)"><font color=blue><b>计算器</b></font></a>&nbsp;&nbsp;<a href="#" onClick="vbscript:ExecShell('D:\\接单系统\\公司简介.exe')"><b>团队</b></a></td>
	  <%
	  	'dim ver
'		ver = conn.execute("select [version] from sysconfig")(0)
'		if ver="Customer" then 
'			response.write "<td colspan=2 width='494'><input type='hidden' name='money' id='money' value=0><input type='hidden' name='savemoney' id='savemoney' value=0></td>"
'		else
	  %>
    <td width="150" align="right" valign="middle" class="font">
	<input name="inp_ordercostcontrol" type="hidden" id="inp_ordercostcontrol" value="<%=OrderCostControl%>">
	<input name="inp_costpoint" type="hidden" id="inp_costpoint" value="<%=orderCostPoint%>">套系金额：</td>
    <%
	dim taoxiprice
	taoxiprice = conn.execute("select money from jixiang where id="&id)(0)
	%>
    <td width="344" class="font"><input name="money" type="text" id="money" size="5" value="<%=taoxiprice%>" />
      元
        <input name="inp_oldprice" type="hidden" id="inp_oldprice" value="<%=taoxiprice%>"> 
        &nbsp; <%if ver<>"Customer" then%>预收金额：
      <input name="savemoney" type="text" id="savemoney" size="5"> 元&nbsp;&nbsp;
 	  <input type="checkbox" name="wzsk" value="yes" > 刷卡收款<%
	else
		response.write "<input type='hidden' name='savemoney' id='savemoney' value=0>"
		response.write "&nbsp;&nbsp;选片后期：<input type='text' name='txt_fujia' id='txt_fujia' size='5' value=0> 元"
	end if
%>        </td>
    <%'end if%>
  </tr>
  <tr align="left" valign="middle" bgcolor="#FFFFFF">
    <td width="150" height="31" align="right" class="font">摄影日期1：</td>
    <td width="344" height="31" class="font"><input name="pz_time" type="text" maxlength="10" id="pz_time" size="13" onChange="CheckDateInfo('pz_time','sp_pz_time')" onBlur="CheckDateInfo('pz_time','sp_pz_time')"  />
        <a onClick="return showCalendar('pz_time', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
        <input name="pz" type="text" size="3">
        点<span id="sp_pz_time" style="padding-left:15px"></span></td>
    <td width="150" align="right" valign="middle" class="font">拍照礼服：</td>
    <td width="344" class="font"><input name="pzlf_time" type="text" maxlength="10" id="pzlf_time" size="13" />
      <a onClick="return showCalendar('pzlf_time', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
      <input name="pzlf" type="text" size="3" value="">
      &nbsp;点 (可为空)</td>
  </tr>
  <tr align="left" valign="middle" bgcolor="#FFFFFF">
    <td height="30" align="right"class="font">摄影日期2：</td>
    <td height="31" class="font"><input name="pz_time2" type="text" maxlength="10" id="pz_time2" size="13" />
      <a onClick="return showCalendar('pz_time2', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
      <input name="pz2" type="text" id="pz2" size="3">
点</td>
    <td height="30" align="right"class="font">结婚礼服：</td>
    <td height="30" class="font"><input name="jhlf_time" type="text" id="jhlf_time" size="13" />
        <a onClick="return showCalendar('jhlf_time', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
        <input name="jhlf" type="text" id="jhlf" size="3">
  &nbsp;点 (可为空)</td>
  </tr>
  <tr align="left" valign="middle" bgcolor="#FFFFFF">
    <td  width="150" height="30" align="right"class="font">选片日期：</td>
    <td width="344" height="31" class="font"><input name="kj_time" type="text" maxlength="10" id="kj_time" size="13" onChange="CheckDateInfo('kj_time','sp_kj_time')"  onBlur="CheckDateInfo('kj_time','sp_kj_time')" />
        <a onClick="return showCalendar('kj_time', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
        <input name="kj" type="text" size="3">&nbsp;点<span id="sp_kj_time" style="padding-left:15px"></span></td>
    <td height="30" align="right" class="font">结婚化妆：</td>
    <td height="30" class="font"><input name="hz_time" type="text" maxlength="10" id="hz_time" size="13" />
        <a onClick="return showCalendar('hz_time', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
        <input name="hz" type="text" size="3" value=""/>
  &nbsp;点 (配套早妆专用) </td>
  </tr>
  <tr align="left" valign="middle" bgcolor="#FFFFFF">
    <td height="30" align="right" class="font">看版日期：</td>
    <td height="30" class="font"><input name="xg_time" type="text" id="xg_time" size="13" />
        <a onClick="return showCalendar('xg_time', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
        <input name="xg" type="text" size="3" value="">
        点</td>
    <td height="30" align="right" class="font">配送结婚：</td>
    <td height="30" class="font"><input name="jhz_style" type="checkbox" id="jhz_style" value="1" <%if instr(jhz_style,"1")>0 then response.write "checked"%>>
收费妆&nbsp;
<input name="jhz_style" type="checkbox" id="jhz_style" value="2" <%if instr(jhz_style,"2")>0 then response.write "checked"%>>
免费妆</td>
  </tr>
  <tr align="left" valign="middle" bgcolor="#ffffff">
    <td height="30" align="right" class="font">取件日期：</td>
    <td height="30" class="font"><input name="qj_time" type="text" id="qj_time" size="13" />
      <a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
      <input name="qj" type="text" size="3" value="">
      点</td>
    <td height="30" align="right" class="font">回婚妆：</td>
    <td height="30" class="font"><label>
      <input name="hhz_time" type="text" maxlength="10" id="hhz_time" size="13" />
      <a onClick="return showCalendar('hhz_time', 'y-mm-dd');" href="javascript:void(0)"><img src="../Image/Button.gif" width="25" height="17" border="0" align="absmiddle" id="IMG2" /></a>
      <input name="hhz" type="text" size="3" value=""/>
&nbsp;点</label></td>
  </tr>
  <tr align="left" valign="middle" bgcolor="#ffffff">
    <td height="30" align="right" class="font">手动单号：</td>
    <td height="30" colspan="3" class="font"><input name="danhao" type="text" id="danhao" size="13" />      
      &nbsp;&nbsp; 毛片回件情况：
      <input name="stated" type="radio" value="1" checked="checked" />
正常
<input type="radio" name="stated" value="2" />
急
<input type="radio" name="stated" value="3" />
特急&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;拍摄多款选
<input name="sl2" type="text" id="sl2" size="7" value="<%=conn.execute("select sl2 from jixiang where id="&id&"")(0)%>" />
张&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr align="left" valign="middle" bgcolor="#ffffff">
    <td height="30" colspan="4" align="left" valign="top" class="font" style="padding-top:10px"><a name="prolist"></a><div class="div_showprice" onClick="javascript:ShowPriceDiv()">显示/隐藏</div>
      <div id="div_body" class="div_list_body"><div id="div_pro_xx" class="div_list_pro">没有选择任何套系产品.</div>
      </div></td>
    </tr>
  <tr align="left" valign="middle" bgcolor="#ffffff">
    <td height="30" align="right" valign="top" class="font">下单备注：</td>
    <td height="30" colspan="3" valign="middle" class="font"><textarea name="beizhu" cols="70" rows="7" id="beizhu"><%=encode2(conn.execute("select beizhu from jixiang where id="&id&"")(0))%>
    </textarea></td>
    </tr>
</table>

<br>
</div>
<div id="div_taoxi">
<table width="975" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" style="display:none">
  <tr>
    <td style="padding:5px"><%=unencode(conn.execute("select txdh from sysconfig")(0))%></td>
  </tr>
</table>
<%
dim typecounter,prochecked
typecounter = 0
set rs2=server.CreateObject("adodb.recordset")
rs2.open "select * from yunyong_type where ishidden=0 order by px asc",conn,1,1%>
<table width="975"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
  <%while not rs2.eof 
  	typecounter = typecounter + 1
  zz=zz
  %>
  <tr id="<%="tr_jxtype_"&typecounter%>">
    <td bgcolor="#fafafa">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr onClick="javascript:showProList(<%=typecounter%>);" style="cursor:pointer">
          <td width="25"><img src="../images/+.gif" name="<%="img_jxtype_"&typecounter%>" width="20" height="20" border="0" id="<%="img_jxtype_"&typecounter%>"></td>
          <td><strong><%=rs2("name")%></strong> </td>
          <td align="right">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr id="<%="tr_prolist_"&typecounter%>" style="display:none">
    <td bgcolor="#FFFFFF" style="padding:5px">
        <%set rs3=server.CreateObject("adodb.recordset")
	rs3.open "select * from yunyong where type_id="&rs2("id")&" and ishidden=0 order by px",conn,1,1
	if not rs3.eof then
		
		%>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
        <%
			do while not rs3.eof
		%>
        <tr onMouseOver="this.bgColor='#FFECFF'" onMouseOut="this.bgColor='#FFFFFF'">
          <%
			for a=1 to 3
				zz=zz+1
				prochecked=false
				if not rs3.eof then
					i=i-1
					if len(zz)=1 then
						zz="00"&zz
					elseif len(zz)=2 then
						zz="0"&zz
					end if
			%>
          <td  align=center valign="top" width="32%" id="<%="td_"&zz%>">
            <table width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
              <tr>
                <td align="center">
                  <div align="left" style="word-space:nowrap"><input type="checkbox" id="check" name="check" value="<%=rs3("id")%>" <%
		  if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then
		  	namelist=namelist&","&rs3("yunyong")       ''''''''''''''''''''''''''
			typelist=typelist&","&rs3("type")       ''''''''''''''''''''''''''
			if pageInvisSetting=1 then 
				xclist=xclist&",0"
			 else
				xclist=xclist&","&rs3("isxc")       ''''''''''''''''''''''''''
			end if
			
			moneylist=moneylist&","&rs3("money")       ''''''''''''''''''''''''''
			costlist=costlist&","&rs3("in_money")       ''''''''''''''''''''''''''
			counterlist=counterlist&","&typecounter       ''''''''''''''''''''''''''
			
			prochecked=true
		  	response.Write "checked"
		  end if
		%> onClick="EditProList(this,'<%=zz%>','<%=rs3("yunyong")%>','<%=rs3("id")%>','<%if pageInvisSetting=1 then 
		response.write "0"
	 else
	 	response.write rs3("isxc")
	end if%>','<%=rs3("type")%>','<%=rs3("money")%>','<%=typecounter%>','<%=rs3("in_money")%>')">
        <%
		if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then
		  	numlist=numlist&","&zz       ''''''''''''''''''''''''''
		  end if
		response.write "<font color='#cccccc'>"&zz&"</font>&nbsp;-&nbsp;"
		dim propic
		propic = GetProFirstPhoto(rs3("id"))
		if not isnull(propic) then
			response.write "<a href='###zoom' onclick=""javascript:zoom(this, '../upload/"&propic&"', '"&rs3("id")&"');"" title='"&rs3("yunyong")&vbcrlf&"价格"&rs3("money")&"元"&vbcrlf&"点击查看套系图片'><font color=red>"&rs3("yunyong")&"</font></a>"
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
		
		if rs3("isxc")=1 and pageInvisSetting=0 then
		%><input name="p<%=rs3("id")%>" type="text" id="p<%=rs3("id")%>" size="1" value="<%=tmp_xc%>" onBlur="EditPageVol(this,'<%=zz%>')">
P&nbsp;<%
		else%>
		<input name="p<%=rs3("id")%>" type="hidden" id="p<%=rs3("id")%>" value="<%=tmp_xc%>">
&nbsp;
		<%end if%><input name="sl<%=rs3("id")%>" type="text" id="sl<%=rs3("id")%>" size="1" value="<% if instr(", "&yunyong11&", ",", "&rs3("id")&", ")>0 then 
		  tt=split(yunyong11,", ")
		  for y=lbound(tt) to ubound(tt)
		  if trim(tt(y))=trim(rs3("id")) then t3=y
		  next 
		  x=split(sl11,", ")
		  response.Write x(t3)
		  sllist=sllist&","&x(t3)       ''''''''''''''''''''''''''
		  end if%>" onBlur="EditProVol(this,'<%=zz%>')">&nbsp;说明<input type="text" name="<%="desc"&rs3("id")%>"id="<%="desc"&rs3("id")%>" size="4"></span></td>
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
<input name="pageInvisSetting" type="hidden" id="pageInvisSetting" value="<%=pageInvisSetting%>">
<input name="newOrderVerify" type="hidden" id="newOrderVerify" value="<%=newOrderVerify%>">
<input name="inp_typecounter" type="hidden" id="inp_typecounter" value="<%=typecounter%>">
<input name="inp_priceflag" type="hidden" id="inp_priceflag" value="">
<br />
<table width="100%" border="0" cellspacing="0" cellpadding="0" id="tb_sumbit">
  <tr>
    <td align="center"><input type="button" name="btn_save" value="保存" style="width:100px" onClick="chk()" />
       <input type="reset" name="Submit" value="重置" style="width:100px" /><br></td>
  </tr>
</table>
<A name="bottom"></A>
<br />
</form>
<script language=javascript>
RefreshCookie("<%=numlist%>","<%=namelist%>","<%=sllist%>","<%=pagelist%>","<%=xclist%>","<%=typelist%>","<%=moneylist%>","<%=counterlist%>","<%=costlist%>");
//InitListBody1();

function GetJstLoginHtml(action, userlevel){
	var _lvtext = "";
	var _sql = "";
	var _htmlstring = "";
	
	var _tmpdept;
	if (userlevel == "") 
		return "";
	else
		_tmpdept = "," + userlevel + ",";

	if (action == "discount") {
		for (var j=1; j<=4; j++){
			if (_tmpdept.indexOf("," + j + ",") < 0 || GetDiscount(j) <= 0)
				continue;

			_lvtext += "、" + GetDeptName(j);
		}
		_lvtext += "、" + GetDeptName(10);
		_lvtext = _lvtext.substring(1);
	}
	else if (action == "costpoint") {
		_lvtext = GetDeptName(2);
	}
	
	_htmlstring += "<table width='90%' border='0' cellspacing='0' cellpadding='0'>";
	_htmlstring += "<tr><td height='60'>";
	_htmlstring += "<table width='100%' border='0' cellspacing='0' cellpadding='0'>";
	_htmlstring += "<tr><td>您必须有" + _lvtext + "访问密码才能修改套系内容或金额。请在下面选择帐户并输入密码以便继续。</td></tr></table>";
	_htmlstring += "<form id='zgfrm' name='zgfrm' method='post' action='' style='display:inline'>";
	_htmlstring += "<fieldset style='padding:5px'><legend>权限验证</legend>";
	_htmlstring += "选择帐户：<select name='zg_msname' id='zg_msname'>";
	
	<%
	Dim arr_levels(5,3)
	Dim li, rsuser, sqluser
	arr_levels(1,1) = "1"
	arr_levels(1,2) = "门市"
	arr_levels(1,3) = "([level]=1 and zhuguan=0)"
	arr_levels(2,1) = "2"
	arr_levels(2,2) = "门市主管"
	arr_levels(2,3) = "([level]=1 and zhuguan=1)"
	arr_levels(3,1) = "3"
	arr_levels(3,2) = "财务"
	arr_levels(3,3) = "([level]=7 and zhuguan=0)"
	arr_levels(4,1) = "4"
	arr_levels(4,2) = "财务主管"
	arr_levels(4,3) = "([level]=7 and zhuguan=1)"
	arr_levels(5,1) = "10"
	arr_levels(5,2) = "经理"
	arr_levels(5,3) = "([level]=10)"

	Set rsuser = Server.CreateObject("ADODB.RECORDSET")
	For li = 1 To UBound(arr_levels,1)
		Response.Write("if (_tmpdept.indexOf("","" + "& arr_levels(li,1) &" + "","") >= 0){") & vbcrlf
		Response.Write("_htmlstring += ""<optgroup label='"& arr_levels(li,2) &"'>"";") & vbcrlf
		sqluser = "select username,peplename from yuangong where "& arr_levels(li,3) &" and isdisabled=0 order by username asc"
		rsuser.open sqluser,conn,1,1
		Do While Not rsuser.eof
			Response.Write("_htmlstring += ""<option value='"& rsuser("username") &"'>"& rsuser("peplename") &"</option>"";") & vbcrlf
			rsuser.movenext
		Loop 
		rsuser.close
		Response.Write("_htmlstring += ""</optgroup>"";}") & vbcrlf
	Next 
	Set rsuser = Nothing 
	%>
	
	_htmlstring += "</select>";
	_htmlstring += "<br />输入密码：<input type='password' name='zg_mspass' id='zg_mspass' /></fieldset>";
	_htmlstring += "<table width='100%' border='0' cellspacing='0' cellpadding='0'>";
	_htmlstring += "<tr><td height='30'><input type='button' name='zg_btnsend' id='zg_btnsend' value=' 提交 ' style='background-color:#efefef' onClick='javascript:CheckZgInfo();' />&nbsp;";
	_htmlstring += "<input type='reset' name='zg_btnreset' id='zg_btnreset' value='重置'  style='background-color:#efefef' /></td>";
	_htmlstring += "</tr></table></form>";
	_htmlstring += "<div id='div_zgmsg'></div>";
	_htmlstring += "</td></tr></table>";
	
	return _htmlstring;
}
</script>

</div>
</body>
</html>