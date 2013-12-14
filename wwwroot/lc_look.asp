<!--#include file="connstr.asp"-->
<!--#incl1ude file="session.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/SystemWorkflow.Class.asp"-->
<%
dim sFromtime,sTotime,sType,sKeyword,sShowflag
sFromtime = request.form("fromtime")
sTotime = request.form("totime")
sType = request.form("fenlei")
sKeyword = trim(request.form("kehu_name"))
sShowflag = request.form("chk_showflag")

dim clsWorkflow, res, currentWork
set clsWorkflow = new SystemWorkflow
clsWorkflow.DBConnection = conn
clsWorkflow.LoadInstance(false)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>客户流程查询</title>
<link href="zxcss.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.STYLE2 {	color: #FF0000;
	font-weight: bold;
}
-->
</style>
<script language="javascript" src="../inc/func.js" type="text/javascript"></script>
<script src="../Js/Calendar.js"></script>
<link href="../Css/calendar-blue.css" rel="stylesheet">
<script type="text/javascript">
function checkform(){
	var fromtime=document.getElementById("fromtime");
	var totime=document.getElementById("totime");
	var kehu_name=document.getElementById("kehu_name");
	var chk_showflag = document.getElementsByName("chk_showflag");
	if((fromtime.value=="" || totime.value=="") && kehu_name.value==""){
		alert("请先选择查询条件。");
		return false;
	} 
	var checked=false
	for(var i=0;i<chk_showflag.length;i++){
		if(chk_showflag[i].checked){
			checked=true;
			break;
		}
	}
	if(!checked){
		alert("请至少选择验件、收/发件或取件其中一项。");
		return false;
	}
	return true;
}
</script>
</head>

<body>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <form action="lc_look.asp?action=look" method="post" name="form1">
    <tr>
      <td height="30">取件时间<span style="display:">
        <input name="fromtime" type="text" id="fromtime" size="10" value="<%=sFromtime%>">
        <span class="font"><a onClick="return showCalendar('fromtime', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" id="IMG2" align="absMiddle" border="0" /></a></span>&nbsp;至
        <input name="totime" type="text" id="totime" size="10" value="<%=sTotime%>">
        <span class="font"><a onClick="return showCalendar('totime', 'y-mm-dd');" href="#"><img src="../Image/Button.gif" id="IMG2" align="absMiddle" border="0" /></a></span></span>&nbsp; 查询条件
        <select name="fenlei" id="fenlei">
          <option value="单号"<%if sType="单号" then response.write " selected"%>>单号</option>
          <option value="姓名"<%if sType="姓名" then response.write " selected"%>>姓名</option>
        </select>
        <input name="kehu_name" type="text" id="kehu_name" value="<%=sKeyword%>" size="18" maxlength="20">
&nbsp;
<input name="chk_showflag" type="checkbox" id="chk_showflag" value="0"<%if instr(sShowflag,"0")>0 then response.write " checked"%>> 
验件
<input name="chk_showflag" type="checkbox" id="chk_showflag" value="1"<%if instr(sShowflag,"1")>0 then response.write " checked"%>> 
收/发件
<input name="chk_showflag" type="checkbox" id="chk_showflag" value="2"<%if instr(sShowflag,"2")>0 then response.write " checked"%>> 
取件
&nbsp;&nbsp;
      <input type="submit" name="Submit" value="查询" onClick="return checkform();"></td>
    </tr>
  </form>
</table>
<%
if (sFromtime<>"" and sTotime<>"") or (sType<>"" and sKeyword<>"") then
	dim sqlstr
	sqlstr = "select k.lxpeple,k.lxpeple2,k.telephone,k.telephone2,s.* from kehu k inner join shejixiadan s on k.id=s.kehu_id where 1=1"
	if sType = "单号" then
		sqlstr = sqlstr & " and (s.danhao='"& sKeyword &"'"
		if isnumeric(sKeyword) then
			sqlstr = sqlstr & " or s.id="& sKeyword
		end if
		sqlstr = sqlstr & ")"
	elseif sType = "姓名" then
		sqlstr = sqlstr & " and (k.lxpeple='"& sKeyword &"' or k.lxpeple2='"& sKeyword &"')"
	else
		response.redirect("lc_look.asp")
		response.end
	end if
	if sFromtime<>"" and sTotime<>"" then
		sqlstr = sqlstr & " and not isnull(qj_time) and qj_time>=#"& sFromtime &"# and qj_time<=#"& sTotime &"# and not isnull(qj_time)"
	end if
	sqlstr = sqlstr & " order by s.times desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sqlstr,conn,1,1
	if rs.eof and rs.bof then 
		response.Write "<br>对不起，没有查询到相关记录!"
		response.End
	else
		dim arr_info(2,1),m,n,tid,tname
		arr_info(0,0)=0
		arr_info(0,1)="验件"
		arr_info(1,0)=5
		arr_info(1,1)="收/发件"
		arr_info(2,0)=1
		arr_info(2,1)="取件"
		while not rs.eof 
%>
<table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#666666">
  <tr bgcolor="#FFFFFF">
    <td style="padding:0 8px 0 8px"><%=rs("id")%>
    <%if rs("danhao")<>"" then response.Write "/"&rs("danhao") %>&nbsp;&nbsp;
    <%response.write rs("lxpeple")
	if rs("telephone")<>"" and not isnull(rs("telephone")) then
		response.write "/"&rs("telephone")
	end if%>
    &nbsp;&nbsp;<%response.write rs("lxpeple2")
	if rs("telephone2")<>"" and not isnull(rs("telephone2")) then
		response.write "/"&rs("telephone2")
	end if%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td style="padding:0 8px 0 8px"><%response.write clsWorkflow.WorkHistory(rs("id"))%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td valign="top" style="padding:10px 5px 5px 5px"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <%
  if sShowflag<>"" then 
	for m=0 to ubound(arr_info,1)
      if instr(sShowflag,cstr(m))>0 then
		tid =arr_info(m,0)
		tname=arr_info(m,1)
		dim rsve
		set rsve=server.CreateObject("adodb.recordset")%>
        <tr><td><b><%=tname%>情况</b>：</td></tr>
		<tr>
          <td height=20>&nbsp; ◇套系取件◇</td>
        </tr>
        <tr>
          <td valign="top"><%
	'套系取件产品组合
	txoldval = ""
	txoldsl = ""
      if isnull(rs("yunyong")) then
        response.Write "没有套系应有!"
      else
        id=split(rs("yunyong"),", ")
        sl=split(rs("sl"),", ")
        if not isnull(rs("wc")) then
            wc=split(rs("wc"),", ")
        end if
      end if
    %>
              <div style="width:97%; padding:3px; margin-left:8px">
                <table width="100%"  border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <%count11=ubound(id)+1
            
            '读取修改记录
            'dim rs_lieyin
            'set rs_lieyin = conn.execute("select * from lieyin where xiangmu_id="&rs("id"))
            counts=0
            for yy=1 to count11
                set rsflag = conn.execute("select [type] from yunyong where id="&id(yy-1))
                if not rsflag.eof and rsflag("type")=1 then
                    counts=counts+1
            %>
                    <td><%
					set rs_yunyong=conn.execute("select id,yunyong from yunyong where id="&id(yy-1)&"")
					response.write "<input type='checkbox' name='chk_txqj' id='chk_txqj' value='"&rs_yunyong("id")&"|"&sl(yy-1)&"'"
					
					rsve.open "SELECT D.ProID, D.ProVol FROM VerifyProDetails D INNER JOIN VerifyProList L ON D.MainID = L.ID WHERE D.ProID="&rs_yunyong("id")&" AND D.Types=0 AND L.vType="&tid&" AND L.Xiangmu_ID="&rs("id"),conn,1,1
					if not (rsve.eof and rsve.bof) then
						txoldval = txoldval & ", " & rs_yunyong("id") & "|" & rsve("ProVol")
						response.write " checked"
					end if
					response.write ">&nbsp;"
					rsve.close
					
					'response.write "<input type='textbox' name='txt_txqj' id='txt_txqj' value='"&sl(yy-1)&"' style='display:none'>"
					response.Write rs_yunyong("yunyong")&"&nbsp;"
					response.Write "- "&sl(yy-1)
					rs_yunyong.close()
					%>                      </td>
                    <%
                if counts mod 4 =0 then response.write "</tr><tr>"
                end if
                rsflag.close()
                next
				
				if txoldval<>"" then
					txoldval = mid(txoldval,3)
				end if
                %>
                  <tr>
                </table>
              </div>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="20">&nbsp;&nbsp;◇选片后期◇</td>
                </tr>
                <tr>
                  <td valign="top" ><div style="width:97%; padding:3px; margin-left:8px">
                      <%
				hqoldval = ""
				hqoldsl = ""
				
				response.Write "<table width=100% border=0 cellspacing=2 cellpadding=0 align=center><tr>"
				set rs2=conn.execute("select fujia.* from fujia inner join yunyong on fujia.jixiang=yunyong.id where yunyong.type=1 and fujia.xiangmu_id="&rs("id")&" order by times")
				'number111=conn.execute("select count(*) from fujia where xiangmu_id="&rs("id")&" and datevalue(times)=#"&rs1("dated")&"# ")(0)
				counter=0
				while not rs2.eof 
				response.write "<td width='50%'>"
				response.write "<input type='checkbox' name='chk_xphq' id='chk_xphq' value='"&rs2("jixiang")&"|"&rs2("sl")&"'"
				
				rsve.open "SELECT D.ProID, D.ProVol FROM VerifyProDetails D INNER JOIN VerifyProList L ON D.MainID = L.ID WHERE D.ProID="&rs2("jixiang")&" AND D.Types=1 AND L.vType="&tid&" AND L.Xiangmu_ID="&rs("id"),conn,1,1
				if not (rsve.eof and rsve.bof) then
					hqoldval = hqoldval & ", " & rs2("jixiang") & "|" & rsve("ProVol")
					response.write " checked"
				end if
				response.write ">&nbsp;"
				rsve.close
				
				'response.write "<input type='textbox' name='txt_xphq' id='txt_xphq' value='"&rs2("sl")&"' style='display:none'>"
				response.Write conn.execute("select yunyong from yunyong where id="&rs2("jixiang"))(0)&"/"&rs2("sl")&"/"
				response.Write rs2("money")&"元&nbsp;&nbsp;说明："&rs2("beizhu")&"&nbsp;&nbsp;"
				response.write "</td>"
				counter=counter+1
				if counter mod 2 = 0 then response.write "</tr><tr>"
				rs2.movenext
				wend 
				if counter mod 2>0 then
					for l = 1 to 2-(counter mod 2)
						response.write "<td width='50%'></td>"
					next
				end if
				rs2.close
				set rs2=nothing
				response.write "</tr></table>"
				
				if hqoldval<>"" then
					hqoldval = mid(hqoldval,3)
				end if
				%>
                    </div>                </tr>
            </table></td>
        </tr>
          <%if m<ubound(arr_info,1) then%>
        <tr><td><hr size="1"></td></tr>
        <%
		end if
	  end if
	next
  end if%>
      </table></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td valign="top" style="padding:5px"><%response.write "摄影"
	if isnull(rs("pz_time")) then
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	else
		response.Write rs("pz_time")
	end if
	response.write "&nbsp;&nbsp;"
	
	response.write "选片"
	if isnull(rs("kj_time")) then
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	else
		response.Write rs("kj_time")
	end if
	response.write "&nbsp;&nbsp;"
	
	response.write "看版"
	if isnull(rs("xg_time")) then
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	else
		response.Write rs("xg_time")
	end if
	response.write "&nbsp;&nbsp;"
	
	response.write "取件"
	if isnull(rs("qj_time")) then
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	else
		response.Write rs("qj_time")
	end if
	response.write "&nbsp;&nbsp;"
	
	response.write "结婚妆"
	if isnull(rs("hz_time")) then
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	else
		response.Write rs("hz_time")
	end if
	response.write "&nbsp;&nbsp;"%>
<strong>总款未缴：<font color=red><%=FinalMoneySum(rs("id"),false)%></font><span class="style1">元</span></strong></td>
</tr>
</table>
<br>
<%rs.movenext
  wend
  rs.close
  set rs=nothing%>
<% 
end if
end if

set clsWorkflow = nothing
%>
</body>
</html>

