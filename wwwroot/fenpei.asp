<!--#include file="connstr.asp"-->
<!--#include file="../inc/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ɴ����ϵͳ -- ����ȷ��</title>
<link href="zxcss.css" rel="stylesheet" type="text/css">
<script src="../Js/Calendar.js"></script>
<script language="javascript" src="../inc/func.js" type="text/javascript"></script>
<script language="javascript">
function elshow(obj){
	var el = document.getElementById(obj);
	var sp = document.getElementById("sp_text");
	if(el.style.display=="none"){
		el.style.display="";
		sp.innerHTML = "������Ŀ"
	}
	else{
		el.style.display="none";
		sp.innerHTML = "��ʾȫ��"
	}
}
function showObjectService(id){
	var counts = document.getElementById("objcounts").value;
	var flag = false;
	for(var i=0;i<counts;i++){
		if(document.getElementById("img_arrow"+i).src.indexOf("arrow_up.jpg")>=0){
			if(parseInt(id)==i) flag = true;
			document.getElementById("img_arrow"+i).src="../Image/arrow_down.jpg";
			document.getElementById("tr_serv"+i).style.display="none";
		}
	}
	if(!flag){
		document.getElementById("tr_serv"+id).style.display="";
		document.getElementById("img_arrow"+id).src="../Image/arrow_up.jpg";
	}
}
function chkfrom(){
	var frm = document.all.form1;
	if(!CheckIsNull(frm.sjs,"��ѡ������Ա��")) return false;
	if(frm.tsVolume)
		if(!CheckIsNumericOrNull(frm.tsVolume,"����д��ɫ������","��ɫ������д����"))��return false;
	if(frm.cpVolume)
		if(!CheckIsNumericOrNull(frm.cpVolume,"����д��Ӱ������","��Ӱ������д����"))��return false;
	if(frm.flag.value=="True"){
		var arr_idlist = frm.id.value.split(", ");
		for(var i=0;i<arr_idlist.length;i++){
			if(document.all("xg_opt"+i).value=='1'){
				if(document.all("xg_time"+i).value==""){
					alert("��ѡ�񿴰�����.");
					document.all("xg_time"+i).focus();
					return false;
				}
				if(document.all("xg"+i).value==""){
					alert("�����뿴��ʱ��.");
					document.all("xg"+i).focus();
					return false;
				}
			}
			if(document.all("sc_time"+i).value==""){
				alert("��ѡ���������.");
				document.all("sc_time"+i).focus();
				return false;
			}
			if(document.all("sc"+i).value==""){
				alert("���������ʱ��.");
				document.all("sc"+i).focus();
				return false;
			}
			
		}
	}
	/*else{
		if(frm.pageid){
			if(frm.pageid.length){
				for(var i=0;i<frm.pageid.length;i++){
					if(!CheckIsNumeric(document.getElementById("p"+frm.pageid[i].value),"P��������Ϊ�ղ���ֻ��������.")) return false;
				}
			}
			else{
				if(!CheckIsNumeric(document.getElementById("p"+document.form1.pageid.value),"P������Ϊ�ղ���ֻ��������.")) return false;
			}
		}
	}*/
	if(frm.xg_time){
		if(frm.xg_opt.value=='1'){
			if(document.all("xg_time").value==""){
				alert("��ѡ�񿴰�����.");
				document.all("xg_time").focus();
				return false;
			}
			if(frm.xg.value==""){
				alert("�����뿴��ʱ��.");
				frm.xg.focus();
				return false;
			}
		}
		if(frm.sc_time.value==""){
			alert("��ѡ���������.");
			frm.sc_time.focus();
			return false;
		}
		if(frm.sc.value==""){
			alert("���������ʱ��.");
			frm.sc.focus();
			return false;
		}
	}
	if(frm.xp2_time){
		if(document.all("xp2_time").value==""){
			alert("��ѡ�����ⷢ����.");
			document.all("xp2_time").focus();
			return false;
		}
		if(frm.xp2.value==""){
			alert("�����뾫���ⷢʱ��.");
			frm.xp2.focus();
			return false;
		}
	}
}

</script>
<link href="../Css/calendar-blue.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE2 {color: #FFCC99}
.style3 {color: #FF0000}
-->
</style>
</head>

<body>
<% level=request("level")
select case session("level")
case 2 
	level=2
	lvname="����ʦ"
case 4
	level=4
	lvname="��Ӱʦ"
case 5 
	level=5
	lvname=GetDutyName(5)
case 10 
	level=10
	lvname="�ܾ���"
case 14 
	level=5
	lvname=GetDutyName(14)
end select

if request("action")="xp2" then level=2

dim id,xg,content
dim arr_id,arr_xg,arr_content
dim hstype,hssql

id = request("id")
xg = request("xg")
content = request("content")

arr_id = split(id,", ")
arr_xg = split(xg,", ")
arr_content = split(content,", ")

if instr(id,",")>0 then flag=true

if request("action2")="edit" then
dim username,userid,username2,userid2
if request("sjs")<>"" then
	username=GetFieldDataBySQL("select peplename from yuangong where username='"&request("sjs")&"'","str","")
	userid=GetFieldDataBySQL("select id from yuangong where username='"&request("sjs")&"'","int",0)
end if
if request("sjs2")<>"" then
	username2=GetFieldDataBySQL("select peplename from yuangong where username='"&request("sjs2")&"'","str","")
	userid2=GetFieldDataBySQL("select id from yuangong where username='"&request("sjs2")&"'","int",0)
end if
select case request("action")
case "ts"
	conn.execute("update  shejixiadan set lc_xp=now,xp_name='"&username&"',tsVolume="&request("tsVolume")&",cpVolume="&request("cpVolume")&" where id="&id&"")
	if username<>session("username") then
		conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&id&",'"&session("userid")&"','[��ɫ]��ǩ��"&lvname&"��"&session("username")&"','������',#"&now()&"#)")
	end if
case "cp"
	cp_wedvol = request("cp_wedvol")
	if cp_wedvol="" or not isnumeric(cp_wedvol) then cp_wedvol=0
	conn.execute("update  shejixiadan set cp_name='"&username&"',cp_memo='"&request("cp_memo")&"',cp_wedvol="&cp_wedvol&" where id="&id)
	if username<>session("username") then
		conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&id&",'"&session("userid")&"','[��Ӱ]��ǩ��"&lvname&"��"&session("username")&"','������',#"&now()&"#)")
	end if
case "hz"
	dim hzexist,rshzexist,existname,hzexist2,existname2
	hzexist = false
	set rshzexist = conn.execute("select hz_name,hz_name2nd from shejixiadan where id="&id)
	if not (rshzexist.eof and rshzexist.bof) then
		if not isnull(rshzexist("hz_name")) and rshzexist("hz_name")<>"" then
			existname=rshzexist("hz_name")
			hzexist=true
		end if
		if not isnull(rshzexist("hz_name2nd")) and rshzexist("hz_name2nd")<>"" then
			existname2=rshzexist("hz_name2nd")
			hzexist2=true
		end if
	end if
	rshzexist.close
	set rshzexist = nothing
	
	set hstype=server.createobject("adodb.recordset")
	hssql = "select * from hs_signtype order by px asc"
	hstype.open hssql,conn,1,1
	
	if username="" and not hzexist then '
		response.write "<script language='javascript'>alert('��ѡ��"&GetDutyName(5)&".');history.back();</script>"
		response.end
	else
		if conn.execute("select count(*) from yuangong where level=14 and isdisabled=0")(0)>0 then
			if request("hz_name2") = "" then
				response.write "<script language='javascript'>alert('��ѡ��"&GetDutyName(14)&".');history.back();</script>"
				response.end
			end if
		end if
		if not hzexist then
			conn.execute("update  shejixiadan set hz_name='"&username&"',lc_hz=now where id="&id&"")
			conn.execute("insert into xiadan (userid,xiangmu_id,type,times) values ('"&request("sjs")&"','"&id&"',5,#"&date&"#)")
			if username<>session("username") then
				conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&id&",'"&session("userid")&"','[����ױ]��ǩ��"&lvname&"��"&session("username")&"','������',#"&now()&"#)")
			end if
			
			if not (hstype.eof and hstype.bof) then
				hstype.movefirst
				do while not hstype.eof
					hssl = request.form("hstype_hzs_"& hstype("id"))
					if hssl<>"" and isnumeric(hssl) then
						conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&id&","&userid&","&hssl&")")
					end if
					hstype.movenext
				loop
			end if
		end if
		if not hzexist2 and username2<>"" then
			conn.execute("update shejixiadan set hz_name2nd='"&username2&"' where id="&id)
			conn.execute("insert into xiadan (userid,xiangmu_id,type,times) values ('"&request("sjs2")&"','"&id&"',5,#"&date&"#)")
			if username2<>session("username") then
				conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&id&",'"&session("userid")&"','[����ױ]��ǩ��"&lvname&"��"&session("username")&"','������',#"&now()&"#)")
			end if
			
			if not (hstype.eof and hstype.bof) then
				hstype.movefirst
				do while not hstype.eof
					hssl = request.form("hstype_hzs2_"& hstype("id"))
					if hssl<>"" and isnumeric(hssl) then
						conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&id&","&userid2&","&hssl&")")
					end if
					hstype.movenext
				loop
			end if
		end if
		
		if request("hz_name2")<>"" then
			conn.execute("update shejixiadan set hz_name2='"&request("hz_name2")&"' where id="&id&"")
			if request("hz_name2")<>session("username") then
				conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&id&",'"&session("userid")&"','[����ױ����]��ǩ��"&lvname&"��"&session("username")&"','������',#"&now()&"#)")
			end if
			
			if not (hstype.eof and hstype.bof) then
				hstype.movefirst
				do while not hstype.eof
					hssl = request.form("hstype_hzzl_"& hstype("id"))
					if hssl<>"" and isnumeric(hssl) then
						conn.execute("insert into hs_signhistory (typeid,xiangmu_id,userid,vol) values ("&hstype("id")&","&id&","& GetFieldDataBySQL("select id from yuangong where peplename='"&request("hz_name2")&"'","int",0) &","&hssl&")")
					end if
					hstype.movenext
				loop
			end if
			'conn.execute("update  shejixiadan set hz_name3='"&request("hz_name2")&"', hs_wc_time=#"&now()&"# where id="&id)
		end if
	end if
	hstype.close
	set hstype = nothing
case "xp" 
	conn.execute("update shejixiadan set xp_name='"&username&"',lc_xp=now where id="&id&"")
case "jx" 
	conn.execute("update shejixiadan set jx_name='"&username&"',lc_jx=now where id="&id&"")
case "xg"
	xp2_time=trim(request.form("xp2_time"))
	xp2=trim(request.form("xp2"))
	if xp2_time="" then 
		xp2_time="null"
		xp2="null"
	else
		xp2_time="#"&xp2_time&"#"
		xp2="'"&xp2&"'"
	end if
	qj_time=trim(request.form("qj_time"))
	qj=trim(request.form("qj"))
	if qj_time="" then 
		qj_time="null"
		qj="null"
	else
		qj_time="#"&qj_time&"#"
		qj="'"&qj&"'"
	end if
	conn.execute("update shejixiadan set xg_sj=now,xg_name='"&username&"',xp2_time="&xp2_time&",xp2="&xp2&",qj_time="&qj_time&",qj="&qj&" where id="&id&"")
case "ky"
	conn.execute("update shejixiadan set lc_ky=now,ky_name='"&username&"' where id="&id&"")
case "xp2"
	conn.execute("update shejixiadan set xp2_name='"&username&"',lc_xp2=now where id="&id)
	'����P��
	set rs2=server.createobject("adodb.recordset")
	rs2.open "select yunyong,pagevol from shejixiadan where id="&id,conn,1,3
	if not (rs2.eof and rs2.bof) then
		if rs2("yunyong")<>"" and not isnull(rs2("yunyong")) then
			pgflag = false
			txyy = rs2("yunyong")
			txpg = rs2("pagevol")
			arr_yy = split(txyy,", ")
			pageid = request.form("pageid_"&id)
			if pageid<>"" then
				arr_pgid = split(pageid,", ")
				for pi = 0 to ubound(arr_yy)
					for pj = 0 to ubound(arr_pgid)
						if cstr(arr_yy(pi))=cstr(arr_pgid(pj)) then
							pgflag = true
							txt_pagevol = trim(request.form("p_"&id&"_"&arr_pgid(pj)))
							if txt_pagevol="" then
								pg_vol = pg_vol & ", 0"
							else
								pg_vol = pg_vol & ", " & txt_pagevol
							end if
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
			rs2.update
		end if
	end if
	rs2.close
	set rs2=nothing
	
	'��������P��
	pageid_fujia = request.form("pageid_fujia_"&id)
	arr_pageid_fujia = split(pageid_fujia,", ")
	for i = 0 to ubound(arr_pageid_fujia)
		pagevol = request.form("p_fujia_"&id&"_"&arr_pageid_fujia(i))
		if pagevol<>"" and isnumeric(pagevol) then
			conn.execute("update fujia set pagevol="&pagevol&" where id="&arr_pageid_fujia(i))
		end if
	next
	
	if session("level")=1 then
		dim altmsg,altcount
		altcount=0
		if FinalMoneySum(id,False)<>0 then 
			altcount = altcount + 1
			altmsg = altcount & "���ͻ���δ����ɿ\n"
		end if
		if NOT CheckTaskEnd(id) then
			altcount = altcount + 1
			altmsg = altcount & "���ͻ���δ��ɵĻ�ױ����Ӱ������\n"
		end if
		
		if altcount>0 then
			altmsg = "�����ⷢȷ�ϳɹ�������������ԭ��δ���ȡ����\t\n"&altmsg
		else
			conn.execute("update shejixiadan set lc_wc=now,wc_name='"&session("username")&"' where id="&id&"")
			altmsg = "�����ⷢ��ȡ��ȷ�ϳɹ���"
		end if
	
		response.Write "<script>alert('"&altmsg&"');window.opener.location.reload();window.close()</script>"
		Response.End
	end if
	
case "sj"
for i = 0 to UBound(arr_id)
	dim pageid,arr_pgid,pi,pj,txyy,txpg,arr_yy,arr_pg,pg_vol,pgflag,txt_pagevol
	beizhu2=conn.execute("select beizhu2 from jixiang where id="&conn.execute("select jixiang from shejixiadan where id="&arr_id(i))(0))(0)
	if beizhu2="" or isnull(beizhu2) then beizhu2=0
	conn.execute("update  shejixiadan set sj_name='"&username&"' where id="&arr_id(i))
	conn.execute("insert into xiadan (userid,xiangmu_id,type,shejichoucheng,beizhu,times) values ('"&request("sjs")&"','"&arr_id(i)&"',2,0,'"&beizhu2&"',#"&date&"#)")
	
	if level=2 then
		if flag then
			if request("xg_opt"&i)="1" then
				conn.execute("update shejixiadan set xg_time=#"&request("xg_time"&i)&"#,xg='"&request("xg"&i)&"' where id="&arr_id(i))
			end if
			if request("sc_time"&i)<>"" and request("sc"&i)<>"" then
				conn.execute("update shejixiadan set sc_time=#"&request("sc_time"&i)&"#,sc='"&request("sc"&i)&"' where id="&arr_id(i))
			end if
			
			'����P��
			pg_vol = ""
			set rs2=server.createobject("adodb.recordset")
			rs2.open "select yunyong,pagevol from shejixiadan where id="&arr_id(i),conn,1,3
			if not (rs2.eof and rs2.bof) then
				if rs2("yunyong")<>"" and not isnull(rs2("yunyong")) then
					pgflag = false
					txyy = rs2("yunyong")
					txpg = rs2("pagevol")
					arr_yy = split(txyy,", ")
					pageid = request.form("pageid_"&arr_id(i))
					if pageid<>"" then
						arr_pgid = split(pageid,", ")
						for pi = 0 to ubound(arr_yy)
							for pj = 0 to ubound(arr_pgid)
								if cstr(arr_yy(pi))=cstr(arr_pgid(pj)) then
									pgflag = true
									txt_pagevol = trim(request.form("p_"&arr_id(i)&"_"&arr_pgid(pj)))
									'response.write "txt_pagevol="&txt_pagevol&"<br>"
									if txt_pagevol="" then
										pg_vol = pg_vol & ", 0"
									else
										pg_vol = pg_vol & ", " & txt_pagevol
									end if
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
					rs2.update
				end if
				rs2.close
				set rs2=nothing
			end if
			
			'��������P��
			pageid_fujia = request.form("pageid_fujia_"&arr_id(i))
			arr_pageid_fujia = split(pageid_fujia,", ")
			for l = 0 to ubound(arr_pageid_fujia)
				pagevol = request.form("p_fujia_"&arr_id(i)&"_"&arr_pageid_fujia(l))
				if pagevol<>"" and isnumeric(pagevol) then
					conn.execute("update fujia set pagevol="&pagevol&" where id="&arr_pageid_fujia(l))
				end if
			next
		else
			if request("xg_opt")="1" then
				conn.execute("update shejixiadan set xg_time=#"&request("xg_time")&"#,xg='"&request("xg")&"' where id="&id)
			end if
			if request("sc_time")<>"" and request("sc")<>"" then
				conn.execute("update shejixiadan set sc_time=#"&request("sc_time")&"#,sc='"&request("sc")&"' where id="&id)
			end if
			
			'����P��
			
			set rs2=server.createobject("adodb.recordset")
			rs2.open "select yunyong,pagevol from shejixiadan where id="&id,conn,1,3
			if not (rs2.eof and rs2.bof) then
				if rs2("yunyong")<>"" and not isnull(rs2("yunyong")) then
					pgflag = false
					txyy = rs2("yunyong")
					txpg = rs2("pagevol")
					arr_yy = split(txyy,", ")
					pageid = request.form("pageid_"&id)
					if pageid<>"" then
						arr_pgid = split(pageid,", ")
						for pi = 0 to ubound(arr_yy)
							for pj = 0 to ubound(arr_pgid)
								if cstr(arr_yy(pi))=cstr(arr_pgid(pj)) then
									pgflag = true
									'response.write "arr_pgid("&pj&")="&arr_pgid(pj)&"<br>"
									txt_pagevol = trim(request.form("p_"&id&"_"&arr_pgid(pj)))
									'response.write "txt_pagevol="&txt_pagevol&"<br>"
									if txt_pagevol="" then
										pg_vol = pg_vol & ", 0"
									else
										pg_vol = pg_vol & ", " & txt_pagevol
									end if
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
					rs2.update
				end if
			end if
			rs2.close
			set rs2=nothing
		end if
		'����
		if trim(arr_content(i))<>"" then
			set rs3=server.CreateObject("adodb.recordset")
			sql3="select * from sjs_baobiao"
			rs3.open sql3,conn,1,3
			rs3.addnew
			rs3("xiangmu_id")=arr_id(i)
			rs3("baobiao")=HTMLEncode2(trim(arr_content(i)))
			rs3("times")=now()
			rs3("userid")=request("sjs")
			rs3("topeple")="������"
			rs3.update
			rs3.close
			set rs3=nothing
		end if	
	end if
next

case "sc"
conn.execute("update  shejixiadan set sc_name='"&username&"',lc_sc=now where id="&id&"")
case "zd"
conn.execute("update  shejixiadan set zd_name='"&username&"',lc_zd=now where id="&id&"")
case "wc"
conn.execute("update  shejixiadan set lc_wc=now,wc_name='"&username&"' where id="&id&"")
conn.execute("update shejixiadan set lc_sj=now where sj_name<>'' and not isnull(sj_name) and isnull(lc_sj) and id="&id)
conn.execute("update shejixiadan set lc_sc=now where sc_name<>'' and not isnull(sc_name) and isnull(lc_sc) and id="&id)
case "hz2"
conn.execute("update  shejixiadan set hz_userid='"&request("sjs")&"',hz_qm_times=#"&now&"# where id="&id&"")
if request("sjs")<>session("userid") then
	conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&id&",'"&session("userid")&"','[���ױ]��ǩ��"&lvname&"��"&session("username")&"','������',#"&now()&"#)")
end if
end select 

	if session("level")<>1 or request("action")<>"xp2" then
		response.Write "<script>alert('�����ɹ���');window.opener.location.reload();window.close()</script>"
		Response.End
	end if
end if

%>
<table width="100%"  border="0" cellpadding="5" cellspacing="0">
<form action="fenpei.asp?action2=edit" name="form1" method="post" onSubmit="return chkfrom()">
  <tr>
    <td valign="top">�����Ŀ:</td>
    <td valign="top">
	  <%
	    dim newidlist
		newidlist = ""
		set rslist = server.CreateObject("adodb.recordset")
		rslist.open "select * from shejixiadan where id in ("&id&")",conn,1,1
		counts = 0
		do while not rslist.eof
		    newidlist = newidlist & ", " & rslist("id")
			if counts =1 then response.write "<span id='sp_idlist'>"' style='display:none'>"
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#eeeeee" style="margin-bottom:5px">
        <tr>
          <td width="68%">&nbsp;<%=rslist("id")%>&nbsp;&nbsp;<%=conn.execute("select jixiang from jixiang where id="&rslist("jixiang"))(0)%> (<%
	  if rslist("xg_opt")=0 then
	  	response.write "�ڲ�����"
	  else
	  	response.write "�ͻ�����"
	  end if
	  %>)
            <input type="hidden" name="xg_opt<%=counts%>" id="xg_opt<%=counts%>" value="<%=rslist("xg_opt")%>"></td>
          <td width="29%" align="right">ȡ��ʱ��
            <%
	response.Write rslist("qj_time")&"&nbsp;&nbsp;"&rslist("qj")
	%>
��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
          <td width="3%"><%
		  if request("action")="sj" then
		  	response.write "<img src='../Image/arrow_down.jpg' name='img_arrow"&counts&"' width='17' height='16' border='0' id='img_arrow"&counts&"' onClick=""showObjectService('"&counts&"')"" style='cursor:hand' title='����鿴��������'>"
		  end if
		  %></td>
        </tr>
        <tr <%if not flag then response.write "style='display:none'"%>>
          <td colspan="2">&nbsp;<span <%if rslist("xg_opt")=0 then response.write "style='display:none'"%>>���˿���
              <input name="xg_time<%=counts%>" type="text" id="xg_time<%=counts%>" size="12"  value="<%=rslist("xg_time")%>">
            <a onClick="return showCalendar('xg_time<%=counts%>', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" name="IMG2" width="25" height="17" border="0" align="absMiddle" id="IMG4" /></a>
            <input name="xg<%=counts%>" type="text" size="2" value="<%if rslist("xg")="" or isnull(rslist("xg")) then
	response.write "0"
else
	response.write rslist("xg")
end if%>">
��&nbsp;&nbsp;&nbsp; </span><span class="font">������(���ƿ���)</span>
<input name="sc_time<%=counts%>" type="text" id="sc_time<%=counts%>" size="12"  value="<%=rslist("sc_time")%>">
<a onClick="return showCalendar('sc_time<%=counts%>', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" name="IMG2" width="25" height="17" border="0" align="absMiddle" id="IMG3" /></a>
<input name="sc<%=counts%>" type="text" size="2" value="<%if rslist("sc")="" or isnull(rslist("sc")) then
	response.write "0"
else
	response.write rslist("sc")
end if%>">
��&nbsp;&nbsp;&nbsp;(��������23�����24�����ϲŲ��ܿ���)<br>
&nbsp;���Ҫ��
<input name="content" type="text" id="content" value="" size="30"></td>
          <td width="3%"></td>
        </tr>
        <%
		if flag and request("action")="sj" and session("zhuguan")=1 then
		%>
        <tr>
          <td colspan="3" style="padding:3px"><b>���P��:</b>
            <%
  dim idlist,sllist,wclist
  if isnull(rslist("yunyong")) then
		response.Write "<br>û����ϵӦ��!"
	else
		idlist=split(rslist("yunyong"),", ")
		
		if not isnull(rslist("wc")) then
			wclist=split(rslist("wc"),", ")
		end if
%>
            <div style="width:98%; padding:5px; border:dashed 1px #999999;">
              ԤԼ����:
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <%
	  dim count11,count22,rslistflag
	  count11=ubound(idlist)+1
	  if rslist("pagevol")<>"" and not isnull(rslist("pagevol")) then
			sllist=split(rslist("pagevol"),", ")
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
		if len(cstr(count22))=2 then
			response.Write "<strong>"&count22&".</strong>"
		else
			response.Write "<strong>0"&count22&"</strong>"
			response.Write "."
		end if
		
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		response.Write "<input type='hidden' id='pageid_"&rslist("id")&"' name='pageid_"&rslist("id")&"' value='"&idlist(yy-1)&"'>"
		response.Write "<input type='text' id='p_"&rslist("id")&"_"&idlist(yy-1)&"' name='p_"&rslist("id")&"_"&idlist(yy-1)&"' value='"
		if rslist("pagevol")<>"" and not isnull(rslist("pagevol")) then
			response.Write sllist(yy-1)
		end if
		rslist_yunyong.close
		response.write "' size='3'> P"
		response.write "</td></tr></table>"
		%></td>
                      <%
				if count22 mod 3 =0 then response.write "</tr><tr>"
			end if
			end if
			rslistflag.close()
		next
		%>
                </table>
              ��������:
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <%
		
		'�������
		count22=0
		set rslist_yunyong=conn.execute("select fujia.*,yunyong.yunyong from fujia inner join yunyong on fujia.jixiang=yunyong.id where yunyong.type=1 and fujia.xiangmu_id="&rslist("id")&" order by times")
		while not rslist_yunyong.eof 
			response.write "<td><table width='85%'  border='0' cellspacing='0' cellpadding='0'><tr><td>"
			count22 = count22 +1
			if len(cstr(count22))=2 then
				response.Write "<strong>"&count22&".</strong>"
			else
				response.Write "<strong>0"&count22&"</strong>"
				response.Write "."
			end if
			
			response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
			response.Write "<input type='hidden' id='pageid_fujia_"&rslist("id")&"' name='pageid_fujia_"&rslist("id")&"' value='"&rslist_yunyong("id")&"'>"
			response.Write "<input type='text' id='p_fujia_"&rslist("id")&"_"&rslist_yunyong("id")&"' name='p_fujia_"&rslist("id")&"_"&rslist_yunyong("id")&"' value='"
			response.Write rslist_yunyong("pagevol")
			response.write "' size='3'> P"
			response.write "</td></tr></table></td>"
			
			if count22 mod 3 =0 then response.write "</tr><tr>"
			rslist_yunyong.movenext
		wend 
		
		rslist_yunyong.close
		set rslist_yunyong=nothing
		%>
          </table>
            </div>
            <%end if%></td></tr>
            <%end if%>
        <tr id="tr_serv<%=counts%>" style='display:none'>
          <td colspan="3" style="padding:3px"><b>ȡ������:</b>
              <%if isnull(rslist("yunyong")) then
		  			response.Write "<br>û����ϵӦ��!"
		  		else
					idlist=split(rslist("yunyong"),", ")
					sllist=split(rslist("sl"),", ")
					if not isnull(rslist("wc")) then
						wclist=split(rslist("wc"),", ")
					end if
	%>
            <div style="width:100%; border:dashed 1px #999999; padding:3px">
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <%count11=ubound(idlist)+1
			
				count22=0
				for yy=1 to count11
					
					set rslistflag = conn.execute("select [type] from yunyong where id="&idlist(yy-1))
					if not rslistflag.eof then
						if rslistflag("type")=1 then
							count22=count22+1
				%>
                  <td><%
					if len(count22)=2 then
						response.Write "<strong>"&count22&".</strong>"
					else
						response.Write "<strong>0"&count22&"</strong>"
						response.Write "."
					end if
		
					dim yyflag,rslist_yunyong
					set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&idlist(yy-1)&"")
					
					response.Write rslist_yunyong("yunyong")&"&nbsp;"
					 response.Write "- "&sllist(yy-1)
					  rslist_yunyong.close()
					%></td>
					<%
							if count22 mod 3 =0 then response.write "</td><tr>"
						end if
						end if
						rslistflag.close()
					next
					%>
			  </table>
            </div>
          <%end if%>
          <b>������Ŀ:</b>
          <%
		  set rshq = conn.execute("select * from fujia where xiangmu_id ="&rslist("id"))
		  if rshq.eof and rshq.bof then
				response.Write "<br>û�к�����Ŀ!"
			else
		  %>
          <div style="width:100%; border:dashed 1px #999999; padding:3px">
            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <%
				count22=0
				do while not rshq.eof
				'for yy=1 to count11
					dim rslist_hq
					set rslist_hq=conn.execute("select id,yunyong from yunyong where id="&rshq("jixiang"))
					if not rslist_hq.eof then
						count22=count22+1
					%>
					<td><%
						if len(count22)=2 then
							response.Write "<strong>"&count22&".</strong>"
						else
							response.Write "<strong>0"&count22&"</strong>"
							response.Write "."
						end if
						
					 	response.Write rslist_hq("yunyong")&"&nbsp;"
					 	response.Write "- "&rshq("sl")
					%></td>
                	<%
					end if
					rslist_hq.close()
					set rslist_hq = nothing
					if count22 mod 3 =0 then response.write "</td><tr>"
					rshq.movenext
				loop
				rshq.close
				set rshq =nothing
				%>
             </table>
          </div>
          <%end if%></td>
          </tr>
      </table>
	  <%	
			rslist.movenext
			counts = counts + 1
			if counts = rslist.recordcount then response.write "</span>"
		loop
		if newidlist<>"" then newidlist = mid(newidlist,3)
		rslist.close
		set rslist = nothing
		%><input type="hidden" id="objcounts" value="<%=counts%>"></td>
  </tr>
  <tr>
    <td width="12%">��ѡ��Ա��:      </td>
    <td width="88%"><select name="sjs" id="sjs" <%
	if level=5 then
		hz_namex=conn.execute("select hz_name from shejixiadan where id="&id)(0)
		if request("action")="hz" then
			if hz_namex<>"" and not isnull(hz_namex) then response.write "disabled"
		end if
	end if
	%>>
      <option value="">��ѡ��</option>
      <%
	  if level = 4 then
		  dim rs_ygtype
		  set rs_ygtype = conn.execute("select * from worktype where [level]=2 or [level]=4 or [level]=12 order by id")
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
	  else
		  set rs=server.CreateObject("adodb.recordset")
		  if request("action")="hz" or request("action")="hz2" then
		  	rs.open "select * from yuangong where ([level]=5 or [level]=14) and isdisabled=0 order by [level]",conn,1,1
		  else
		  	rs.open "select * from yuangong where [level]="&level&" and isdisabled=0",conn,1,1
		  end if
		  while not rs.eof%>
		  <option value="<%=rs("username")%>" <%
		  	if level=5 then
		  		if rs("peplename")=session("username") or rs("peplename")=hz_namex then response.write "selected"
			else
				if rs("peplename")=session("username") then response.write "selected"
			end if%>><%=rs("peplename")%></option>
		  <%rs.movenext
		  wend
		  rs.close
		  set rs=nothing
	  end if
	 %>
    </select>  
	<%if request("action")="hz" then
		response.write "&nbsp;&nbsp;"
		response.write ShowWedSignInput("hstype_hzs_", id, hz_namex, true)
		response.write "<br />"%>
	<select name="sjs2" id="sjs2" <%
		hz_namex=conn.execute("select hz_name2nd from shejixiadan where id="&id)(0)
		if request("action")="hz" then
			if hz_namex<>"" and not isnull(hz_namex) then response.write "disabled"
		end if
	%>>
      <option value="">��ѡ��</option>
      <%set rs=server.createobject("adodb.recordset")
	  	rs.open "select * from yuangong where ([level]=5 or [level]=14) and isdisabled=0 order by [level]",conn,1,1
		while not rs.eof%>
		  <option value="<%=rs("username")%>" <%
		  	if rs("peplename")=hz_namex then 
				response.write "selected"
			end if%>><%=rs("peplename")%></option>
		  <%rs.movenext
		  wend
		  rs.close
		  set rs=nothing
	 %>
    </select>  
	<%	response.write "&nbsp;&nbsp;"
		response.write ShowWedSignInput("hstype_hzs2_", id, hz_namex, true)
		end if%>&nbsp;&nbsp;<%if request("action")="cp" then%>��Ƭ
      <input name="cp_wedvol" type="text" id="cp_wedvol" value="0" size="5">
      <%
end if
if not flag then
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from shejixiadan where id="&id,conn,1,1
	if request("action")="sj" then  'session("zhuguan")=1 and 
	%>
      <span <%if rs("xg_opt")=0 then response.write "style='display:none'"%>>(
      <%
	  if rs("xg_opt")=0 then
	  	response.write "�ڲ�����"
	  else
	  	response.write "�ͻ�����"
	  end if
	  %>
      <input type="hidden" name="xg_opt" id="xg_opt" value="<%=rs("xg_opt")%>">
  )���˿���
  <input name="xg_time" type="text" id="xg_time" size="11"  value="<%=rs("xg_time")%>">
  <a onClick="return showCalendar('xg_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" name="IMG2" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
  <input name="xg" type="text" size="2" value="<%if rs("xg")="" or isnull(rs("xg")) then
	response.write "0"
else
	response.write rs("xg")
end if%>">
      ��   </span><span class="font">(���ƿ���)</span>
      <input name="sc_time" type="text" id="sc_time" size="8"  value="<%=rs("sc_time")%>">
      <a onClick="return showCalendar('sc_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" name="IMG2" width="25" height="17" border="0" align="absMiddle" id="IMG" /></a>
      &nbsp;
      <input name="sc" type="text" id="sc" value="<%if rs("sc")="" or isnull(rs("sc")) then
	response.write "0"
else
	response.write rs("sc")
end if%>" size="2">
��(�����23�����24�ŲŲ��ܿ���)
      <%end if
end if
if request("action")="ts" then
	  %>
     &nbsp;&nbsp; ��Ƭ����
      <input name="tsVolume" type="text" id="tsVolume" size="10">
     &nbsp;&nbsp;&nbsp; ��Ӱ����
      <input name="cpVolume" type="text" id="cpVolume" size="10">
      <%end if
	  if request("action")="xg" then%>
	  &nbsp;&nbsp;�����ⷢ
	  <input name="xp2_time" type="text" id="xp2_time" size="11"  value="<%=rs("xp2_time")%>">
	  <a onClick="return showCalendar('xp2_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" name="IMG2" width="25" height="17" border="0" align="absMiddle" id="IMG2" /></a>
	  <input name="xp2" type="text" size="2" value="<%if rs("xp2")="" or isnull(rs("xp2")) then
		response.write "0"
	  else
		response.write rs("xp2")
	  end if%>">
      ��  &nbsp; ȡ��ʱ��
	  <input name="qj_time" type="text" id="qj_time" size="11"  value="<%=rs("qj_time")%>">
      <a onClick="return showCalendar('qj_time', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" name="IMG2" width="25" height="17" border="0" align="absMiddle" id="IMG5" /></a>
      <input name="qj" type="text" size="2" value="<%if rs("qj")="" or isnull(rs("qj")) then
		response.write "0"
	  else
		response.write rs("qj")
	  end if%>">
��
<%end if%>
      <input name="action" type="hidden" id="action" value="<%=request("action")%>">
  <input name="id" type="hidden" id="id" value="<%=newidlist%>"></td>
  </tr>
  <%if request("action")="hz" then%>
  <tr>
      <td>��ѡ������:</td>
      <td><select name="hz_name2" id="hz_name2" <%
	  if rs("hz_name2")<>"" and not isnull(rs("hz_name2")) then
	  	response.write "disabled"
	  end if
	  %>>
        <option value="">��ѡ��</option>
        <%set rsyg=server.CreateObject("adodb.recordset")
	  rsyg.open "select * from yuangong where level=14 and isdisabled=0",conn,1,1
	  while not rsyg.eof%>
        <option value="<%=rsyg("peplename")%>" <%if rsyg("peplename")=rs("hz_name2") or rs("hz_name2")=rsyg("peplename") then response.write "selected"%>><%=rsyg("peplename")%></option>
        <%rsyg.movenext
	  wend
	  rsyg.close
	  set rsyg=nothing%>
      </select>
	  <%
	  	response.write "&nbsp;&nbsp;"
		set hstype=server.createobject("adodb.recordset")
		hssql = "select * from hs_signtype order by px asc"
		hstype.open hssql,conn,1,1
		do while not hstype.eof
			response.write hstype("title") & "&nbsp;" & "<input type='text' name='hstype_hzzl_"&hstype("id")&"' size='3'"
			if not isnull(rs("hz_name2")) and rs("hz_name2")<>"" then
				vol=GetFieldDataBySQL("SELECT hs_signhistory.vol FROM hs_signhistory INNER JOIN yuangong ON hs_signhistory.userid = yuangong.ID where yuangong.peplename='"&rs("hz_name2")&"' and hs_signhistory.xiangmu_id="&id&" and hs_signhistory.typeid="& hstype("id"),"int",0)
				response.write " value='"&vol&"' readonly"
			end if
			response.write " />&nbsp;&nbsp;&nbsp;"
			hstype.movenext
		loop
		hstype.close
		set hstype = nothing
	  %>
	  </td>
    </tr>
	<%
	end if
	sjPageInvis = conn.execute("select sjPageInvis from sysconfig")(0)
	if not flag then
		if (request("action")="sj") or (request("action")="xp2" and sjPageInvis=0) then
		' and session("zhuguan")=1
	%>
    <tr>
    <td>���P��</td>
    <td><%
  if isnull(rs("yunyong")) then
		response.Write "<br>û����ϵӦ��!"
	else
		idlist=split(rs("yunyong"),", ")
		
		if not isnull(rs("wc")) then
			wclist=split(rs("wc"),", ")
		end if
%>
      <div style="width:98%; padding:5px; border:dashed 1px #999999;">
      ԤԼ����:
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <%
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
		
		set rslist_yunyong=conn.execute("select id,yunyong from yunyong where id="&idlist(yy-1)&"")
		response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
		response.Write "<input type='hidden' id='pageid_"&rs("id")&"' name='pageid_"&rs("id")&"' value='"&idlist(yy-1)&"'>"
		response.Write "<input type='text' id='p_"&rs("id")&"_"&idlist(yy-1)&"' name='p_"&rs("id")&"_"&idlist(yy-1)&"' value='"
		if rs("pagevol")<>"" and not isnull(rs("pagevol")) then
			response.Write sllist(yy-1)
		end if
		response.write "' size='3'> P"
		rslist_yunyong.close()
		response.write "</td></tr></table>"
		%></td>
            <%
				if count22 mod 3 =0 then response.write "</td><tr>"
			end if
			end if
			rslistflag.close()
		next
		%>
              </table>
              ��������:
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <%
		
		'�������
		count22=0
		set rslist_yunyong=conn.execute("select fujia.*,yunyong.yunyong from fujia inner join yunyong on fujia.jixiang=yunyong.id where yunyong.isxc=1 and fujia.xiangmu_id="&rs("id")&" order by times")
		while not rslist_yunyong.eof 
			response.write "<td><table width='85%'  border='0' cellspacing='0' cellpadding='0'><tr><td>"
			count22 = count22 +1
			if len(cstr(count22))=2 then
				response.Write "<strong>"&count22&".</strong>"
			else
				response.Write "<strong>0"&count22&"</strong>"
				response.Write "."
			end if
			
			response.Write rslist_yunyong("yunyong")&"</td><td align=right>"
			response.Write "<input type='hidden' id='pageid_fujia_"&rs("id")&"' name='pageid_fujia_"&rs("id")&"' value='"&rslist_yunyong("id")&"'>"
			response.Write "<input type='text' id='p_fujia_"&rs("id")&"_"&rslist_yunyong("id")&"' name='p_fujia_"&rs("id")&"_"&rslist_yunyong("id")&"' value='"
			response.Write rslist_yunyong("pagevol")
			response.write "' size='3'> P"
			response.write "</td></tr></table></td>"
			
			if count22 mod 3 =0 then response.write "</tr><tr>"
			rslist_yunyong.movenext
		wend 
		
		rslist_yunyong.close
		set rslist_yunyong=nothing
		%>
          </table>
      </div>
      <%end if%></td>
  </tr>
  <%end if%>
  <tr>
    <td colspan="2">����ȡ��ʱ��
      <%
	response.Write rs("qj_time")&"&nbsp;&nbsp;"&rs("qj")
	%>
��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ǰ8�춨����ʱ��</td>
  </tr>
  
	<%
	rs.close
	set rs=nothing
end if
%><tr <%if request("action")<>"cp" then response.Write "style='display:none'"%>>
    <td>��Ӱ˵��:</td>
    <td><input name="cp_memo" type="text" id="cp_memo" value="" size="50">
      <br>
      �������ɴ����װ����ɫ�������ڡ����⡣</td>
  </tr>
  <tr <%if flag then response.write "style='display:none'"%>>
    <td valign="top"><%if level=5 or level=1 then
		response.write "˵��:"
	else
		response.write "���Ҫ��:"
	end if
	%></td>
    <td><textarea name="content" cols="70" rows="5" id="content"></textarea></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input type="submit" name="Submit" value="ȷ��">
      <input type="reset" name="Submit2" value="����">
      <input name="flag" type="hidden" id="flag" value="<%=flag%>"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p>&nbsp;</p>
      <p>�����볣����<br>
        �ļ���    ���۾�  �޸첲  �����͡����ݡ�������   �����ް׵�  ��С�ۡ��޶��ӡ���ͷ�����Ρ����������� <br>
        ���⾰����    ���Ҫȥ��   �������   ������  ���ֲ�Ҫ   ���⾰��ƣ���ѩ�������ݵء�����  �ٲ����������Դ�<br>
        <br>
        ��������ȡ�����Ӽ�</p>
      <p>��桡�����桡�����Ҫ��<br>
      </p>
      <p> ����Ӱ������<br>
        ���˷�Ӧ  ȡ������ ��ɫ��̫��  �ƹ�����   �ӰѾ�,����������������ĺ���<br>
      </p>
      <p> �����г�����<br>
        ���Ӽ�����, ̫�����   �������ո���    ������������������  �� ����˹�ͨ�����õ㼼�� <br>
      </p>
      <p>&nbsp;&nbsp;&nbsp;<span class="STYLE2">��������ﲻ��,��д���ı�������˾������,���ǽ������Ǽ��� </span></p></td>
  </tr>
  </form>
</table>
</body>
</html>


