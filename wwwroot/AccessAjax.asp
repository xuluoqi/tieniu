<!--#include file="ZLSDK.asp"-->
<!--#include file="connstr.asp"-->
<!--#include file="../inc/sms_class.asp"-->
<!--#include file="../inc/function.asp"-->
<%
response.Charset="gb2312"
dim action
action = Request.QueryString("action")
dim sql,rs,str,i
select case action
	case "checkauthor"
		dim un,pw
		un = trim(request.QueryString("un"))
		pw = trim(request.QueryString("pw"))
		set rs = server.CreateObject("adodb.recordset")
		sql = "select * from sysconfig"
		rs.open sql,conn,1,1
		if (rs("authorUsername")="" or isnull(rs("authorUsername"))) and (rs("authorPassword")="" or isnull(rs("authorPassword"))) then
			str = "0"
		elseif rs("authorUsername") = un and rs("authorPassword") = md5(pw) then
			str = "1"
		else
			str = "-1"
			if un="" and pw="" then str="null"
		end if
		response.write str
		rs.close
		set rs = nothing
		
	case "senddayreport"
		dim str_date
		str_date = cdate(trim(request.QueryString("dates")))
		
	  	set rs = server.CreateObject("adodb.recordset")
		
		dim dd_all_jixiang,dd_all_dingjin,jx_save,per_xmid
		dd_all_jixiang=0
	  	dd_all_dingjin=0
		per_xmid=0

		Dim rsmoney
		set rsmoney=conn.execute("select id,xiangmu_id,money from save_money where xiangmu_id in (select id from shejixiadan where not isnull(times) and datevalue(times)=#"&str_date&"# and not isnull(times)) and not isnull(times) and datevalue(times)=#"&str_date&"# and not isnull(times) and type=1 order by id")
		do while not rsmoney.eof
			if per_xmid<>rsmoney("xiangmu_id") then
				per_xmid=rsmoney("xiangmu_id")
				if rsmoney("money")<>"" and isnumeric(rsmoney("money")) then
					dd_all_dingjin=dd_all_dingjin+rsmoney("money")
				end if
			end if
			rsmoney.movenext
		loop
		rsmoney.close()
		set rsmoney=nothing
		
		dd_all_jixiang = conn.execute("select sum(money) from save_money where type=1 and not isnull(times) and datevalue(times)=#"&str_date&"# and not isnull(times) ")(0)
		if isnull(dd_all_jixiang) then dd_all_jixiang=0
		
		dim hq_all_fujia,hq_all_save,hq_save,hq_money
		hq_all_fujia=0
	  	hq_all_save=0
	  	rs.open "select id,jixiang_money,ky_name from shejixiadan where not isnull(kj_time) and datevalue(kj_time)=#"&str_date&"# and not isnull(kj_time)",conn,1,1
		do while not rs.eof
			hq_money=conn.execute("select sum(money) from fujia where xiangmu_id="&rs("id"))(0)
			if isnull(hq_money) then hq_money=0
			hq_all_fujia = hq_all_fujia + hq_money
			rs.movenext
		loop
	  	rs.close

		hq_all_save=conn.execute("select sum(money) from save_money where not isnull(times) and datevalue(times)=#"&str_date&"# and not isnull(times) and [type]=2")(0)
		if isnull(hq_all_save) then hq_all_save=0
		
		dim pz_all_fujia,pz_all_save,pz_save,pz_money
		pz_all_fujia=0
	  	pz_all_save=0
	  	rs.open "SELECT * FROM shejixiadan Where not isnull(pz_time) and datevalue(pz_time)=#"&str_date&"# and not isnull(pz_time)",conn,1,1
		do while not rs.eof
			pz_money=conn.execute("select sum(money) from fujia2 where xiangmu_id="&rs("id"))(0)
			if isnull(pz_money) then pz_money=0
			pz_all_fujia = pz_all_fujia + pz_money
			rs.movenext
		loop
	  	rs.close

		pz_all_save=conn.execute("select sum(money) from save_money where not isnull(times) and datevalue(times)=#"&str_date&"# and not isnull(times) and [type]=3")(0)
		if isnull(pz_all_save) then pz_all_save=0
		
		dim jh_all_goumai,jh_all_save,jh_money,jh_save
		jh_all_goumai=0
	  	jh_all_save=0
	  	rs.open "SELECT * FROM shejixiadan Where not isnull(hz_time) and datevalue(hz_time)=#"&str_date&"# and not isnull(hz_time)",conn,1,1
		do while not rs.eof
			jh_money=conn.execute("select sum(money) from goumai where xiangmu_id="&rs("id"))(0)
			if isnull(jh_money) then jh_money=0
			jh_all_goumai = jh_all_goumai + jh_money
			rs.movenext
		loop
	  	rs.close
		set rs = Nothing
		
		jh_all_save=conn.execute("select sum(money) from save_money where not isnull(times) and datevalue(times)=#"&str_date&"# and not isnull(times) and [type]=4")(0)
		if isnull(jh_all_save) then jh_all_save=0
		
		dim goumai_money
		goumai_money=conn.execute("select sum(money) from goumai_jilu where not isnull(times) and datevalue(times)=#"&str_date&"# and not isnull(times)")(0)
		if isnull(goumai_money) then goumai_money=0
		
		dim ls_zhichu
		ls_zhichu=conn.execute("select sum(money) from zhichu_jilu where not isnull(times) and datevalue(times)=#"&str_date&"# and not isnull(times) and changshang_id=0")(0)
		if isnull(ls_zhichu) then ls_zhichu=0
		
		dim cjsl,kssl
		cjsl = conn.execute("select count(*) from shejixiadan where datevalue(times)=#"&str_date&"#")(0)
		if isnull(cjsl) then cjsl=0
		kssl = conn.execute("select count(*) from kehu where datevalue(times)=#"&str_date&"# and islost=1")(0)
		if isnull(kssl) then kssl=0
		
		dim cpsl
		cpsl = conn.execute("select count(*) from shejixiadan where not isnull(pz_time) and datevalue(pz_time)=#"&str_date&"# and not isnull(pz_time)")(0)
		if isnull(cpsl) then cpsl=0
		
		dim kysl
		kysl = conn.execute("select count(*) from shejixiadan where not isnull(kj_time) and datevalue(kj_time)=#"&str_date&"# and not isnull(kj_time)")(0)
		if isnull(kysl) then kysl=0
		
		dim sms_message
		sms_message = month(str_date) & "月" & day(str_date)
		sms_message = sms_message & "总收入" & clng(dd_all_jixiang+hq_all_save+pz_all_save+jh_all_save+goumai_money)
		sms_message = sms_message & "定金" & clng(dd_all_dingjin)
		sms_message = sms_message & "拍照" & clng(dd_all_jixiang-dd_all_dingjin)
		sms_message = sms_message & "选片" & clng(hq_all_save)
		sms_message = sms_message & "拍照妆" & clng(pz_all_save)
		sms_message = sms_message & "结婚妆" & clng(jh_all_save)
		sms_message = sms_message & "零散" & clng(goumai_money)
		sms_message = sms_message & "支出" & clng(ls_zhichu)
		sms_message = sms_message & "成交" & cjsl
		sms_message = sms_message & "客失" & kssl
		sms_message = sms_message & "摄" & cpsl
		sms_message = sms_message & "选" & kysl
		
		dim arr_mes()
		if len(sms_message)>70 then
			redim arr_mes(1)
			arr_mes(0) = left(sms_message,70)
			arr_mes(1) = mid(sms_message,71)
		else
			redim arr_mes(0)
			arr_mes(0) = sms_message
		end if
			
		response.write sms_message
		'response.End
		
'		dim rsdx,SMS_id,SMS_sn,SMS_pw,manager_phone
'		Set rsdx = conn.execute("select * from duanxin where statu=1")
'		If Not rsdx.eof Then
'			SMS_id = rsdx("id")
'			SMS_sn = rsdx("zhuce_id")
'			SMS_pw = rsdx("pass_word")
'			manager_phone = rsdx("manager_phone")
'		End If
'		set rsdx = nothing
'		
'		if SMS_sn<>"" and not isnull(SMS_sn) and manager_phone<>"" and not isnull(manager_phone) then
'			dim SMS_Object
'			Set SMS_Object = new SMS_Class
'			SMS_Object.SmsCompanyID = SMS_id
'			SMS_Object.Create()
'			
'			dim ii,re
'			dim ArrList,bound
'			for ii = 0 to ubound(arr_mes)
'				re=SMS_Object.SendMessage(SMS_sn, SMS_pw, manager_phone, arr_mes(ii), "", "")
'				If re=1 then
'					ArrList = Split(manager_phone,",")
'					bound = ubound(ArrList)
'					conn.execute("insert into SmsHistory (UserName,SmsType,SendTime,SmsVolume,Content1,Content2,Content3,Content4) values ('"&session("username")&"',0,#"&now()&"#,"&bound+1&",'"&arr_mes(ii)&"','','','')")
'				Else
'					Response.Write "0"
'					Response.End
'				End If
'			next
'			Response.Write "1"
'		else
'			Response.Write "0"
'		end if
'		set SMS_Object = nothing
		
	case "checkmszg"
		dim zg_name,zg_pass,zg_sql,zg_rs
		zg_name = request.QueryString("zg_name")
		zg_pass = request.QueryString("zg_pass")
		set zg_rs=server.createobject("adodb.recordset")
		zg_sql="select * from yuangong where username='"&zg_name&"'"
		zg_rs.open zg_sql,connstr,1,1
		if (zg_rs.eof and zg_rs.bof) then
			Response.Write "0"
		else
			if Trim(zg_rs("password"))<>md5(zg_pass) then 
				Response.Write "0"
			else
				session("zg_adminid")=zg_rs("id")
				Response.Write "1"
			end if
		end if
		zg_rs.close
		set zg_rs = nothing
	case "checkhqtuser"
		dim hqt_user,hqt_pass
		hqt_user = request.QueryString("hqt_user")
		hqt_pass = request.QueryString("hqt_pass")
		set hqt_rs=server.createobject("adodb.recordset")
		hqt_sql="select id from kehu where hqt_username='"&hqt_user&"'"
		hqt_rs.open hqt_sql,connstr,1,1
		if (hqt_rs.eof and hqt_rs.bof) then
			Response.Write "0"
		else
			Response.Write "1"
		end if
		hqt_rs.close
		set hqt_rs = nothing
	case "checkEnrolProExist"
		dim id,ProID,ProName,ProVol,ProMemo,JxFlag,HqFlag,s,yi
		dim ArrEnrolProID,ArrEnrolProVol,ArrEnrolProMemo
		'id = Request.QueryString("id")
		ProID = Request.QueryString("pid")
		ProVol = Request.QueryString("pvol")
		ProMemo = Request.QueryString("pmemo")
		
		Dim ArrXiangmuID,ChkXmProExists,OrderID,ArrYunyong,ArrProVol
		Dim ProJxVol,ProHqVol
		Dim RsXiangmu,rsyjqj,RsVerify,RsFujia
		Dim ReturnString,ExistProID
		
		If ProID="" Then Response.End()
		ArrEnrolProID = split(ProID,",")
		ArrEnrolProVol = split(ProVol,",")
		ArrEnrolProMemo = split(ProMemo,"|")
		
		for s = 0 to ubound(ArrEnrolProID)
			ExistProID = "":ProName = ""
			ProJxVol = 0:ProHqVol = 0
			dim rstemp
			set rstemp = conn.execute("select yunyong from yunyong where id="&ArrEnrolProID(s))
			if not (rstemp.eof and rstemp.bof) then ProName = rstemp(0)
			rstemp.close
			set rstemp = nothing
			
			if trim(ArrEnrolProMemo(s))<>"" then
				ArrXiangmuID = Split(ArrEnrolProMemo(s),",")
				For i = 0 to UBound(ArrXiangmuID)
					JxFlag = null
					HqFlag = null
					
					ArrXiangmuID(i) = Trim(ArrXiangmuID(i))
					If ArrXiangmuID(i)<>"" And IsNumeric(ArrXiangmuID(i)) Then
						'检查套系内容
						Set RsXiangmu = Server.CreateObject("ADODB.RECORDSET")
						RsXiangmu.open "select yunyong,sl from shejixiadan where instr(', '+yunyong+',', ', "&ArrEnrolProID(s)&",')>0 and id="&ArrXiangmuID(i),conn,1,1
						if Not (RsXiangmu.Eof And RsXiangmu.Bof) Then
							'计算数量
							ArrYunyong = Split(RsXiangmu("yunyong"),", ")
							ArrProVol = Split(RsXiangmu("sl"),", ")
							For yi = 0 To UBound(ArrYunyong)
								If CInt(ArrYunyong(yi)) = CInt(ArrEnrolProID(s)) Then
									ProJxVol = ProJxVol + CInt(ArrProVol(yi))
								End If 
							Next 

							Set RsVerify = Server.CreateObject("ADODB.RECORDSET")
							RsVerify.Open "select d.* from VerifyProDetails d inner join VerifyProList o on d.mainid=o.id where (o.vType=2 or o.vType=0) and d.Types=0 and o.Xiangmu_ID="&ArrXiangmuID(i)&" and d.proid="&ArrEnrolProID(s),conn,1,3
							if not (RsVerify.eof and RsVerify.bof) Then
								JxFlag = 1
							else
								JxFlag = 0
							end if
							RsVerify.close
							set RsVerify=Nothing
						else
							JxFlag = -1
						end If
						RsXiangmu.close
						Set RsXiangmu = Nothing
		
						'检查后期内容
						Set RsFujia = Server.CreateObject("ADODB.RECORDSET")
						RsFujia.open "select fujia.* from fujia inner join yunyong on fujia.jixiang=yunyong.id where yunyong.type=1 and fujia.jixiang="&ArrEnrolProID(s)&" and fujia.xiangmu_id="&ArrXiangmuID(i)&" order by times",conn,1,1
						If Not (RsFujia.Eof And RsFujia.Bof) Then
							'计算后期数量
							ProHqVol = ProHqVol + RsFujia("sl")

							Set RsVerify = Server.CreateObject("ADODB.RECORDSET")
							RsVerify.Open "select d.* from VerifyProDetails d inner join VerifyProList o on d.mainid=o.id where o.vType=2 and d.Types=1 and o.Xiangmu_ID="&ArrXiangmuID(i)&" and d.proid="&ArrEnrolProID(s),conn,1,3
							if not (RsVerify.eof and RsVerify.bof) Then
								HqFlag = 1
							Else
								HqFlag = 0
							end if
							RsVerify.close
							set RsVerify=Nothing
						Else
							HqFlag = -1
						End If 
						RsFujia.close
						Set RsFujia = Nothing 
					End If
					if abs(JxFlag+HqFlag)=2 or (JxFlag+HqFlag=0 and JxFlag<>0) Or ((ProJxVol+ProHqVol<CInt(ArrEnrolProVol(s))) And i=UBound(ArrXiangmuID)) then
						ExistProID = ExistProID & "," & ArrXiangmuID(i)
					end if
				Next
			end If
			if ExistProID<>"" Or ProJxVol+ProHqVol<CInt(ArrEnrolProVol(s)) then
				ExistProID = mid(ExistProID,2)
				ReturnString = ReturnString & "||" & s & "|" & ProName & "|" & ExistProID & "|" & ProJxVol+ProHqVol & "," & CInt(ArrEnrolProVol(s))
			end if
		next
		if ReturnString <> "" then ReturnString = mid(ReturnString,3)
		Response.Write ReturnString
	Case "checkNewOrderPrice"
		Dim str_yunyong,str_volume,price,acc,currentProCost,allProCost
		Dim arr_price_yunyong,arr_price_volume
		str_yunyong = Request.Querystring("yunyong")
		str_volume = Request.Querystring("sl")
		price = Request.Querystring("price")
		allProCost = 0
		If str_yunyong="" Or Not IsNumeric(price) Then
			acc = 1
		Else
			Dim SysconfigPriceBasePoint
			SysconfigPriceBasePoint = conn.execute("select OrderCostPoint from sysconfig")(0)
			If SysconfigPriceBasePoint > 0 Then 
				arr_price_yunyong = Split(str_yunyong,",")
				arr_price_volume = Split(str_volume,",")
				For i = 0 To UBound(arr_price_volume)
					If arr_price_volume(i) <> "" And IsNumeric(arr_price_volume(i)) then
						currentProCost = conn.execute("select in_money from yunyong where yunyong='"&arr_price_yunyong(i)&"'")(0)
						allProCost = allProCost + (currentProCost * CInt(arr_price_volume(i)))
						'response.write "select in_money from yunyong where yunyong='"&arr_price_yunyong(i)&"'"&vbcrlf
						'response.write arr_price_yunyong(i)&"="&currentProCost&vbcrlf
					End If 
				Next 
				If allProCost>0 then
					If (price / allProCost) < (SysconfigPriceBasePoint / 100) Then
						acc = 1
					Else
						acc = 0
					End If 
				Else
					acc = 0
				End If 
			Else
				acc = 0
			End If 
		End If
		response.write acc
		'response.write "acc="&acc&"<br>price="&price&"<br>allProCost="&allProCost&"<br>SysconfigPriceBasePoint="&SysconfigPriceBasePoint
end select
%>