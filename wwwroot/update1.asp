<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
Dim ProjID
ProjID = request.querystring("projid")
If ProjID<>"" And IsNumeric(ProjID) Then
Dim CusID
CusID = conn.execute("select kehu_id from shejixiadan where id="& ProjID)(0)
conn.execute("delete from cuenchu where xiangmu_id="& ProjID)
conn.execute("delete from xiadan where xiangmu_id="& ProjID)
conn.execute("delete from sjs_baobiao where xiangmu_id="& ProjID)
conn.execute("delete from save_money where xiangmu_id="& ProjID)
conn.execute("delete from fujia where xiangmu_id="& ProjID)
conn.execute("delete from fujia2 where xiangmu_id="& ProjID)
conn.execute("delete from goumai where xiangmu_id="& ProjID)
conn.execute("delete from shejixiadan where id="& ProjID)
conn.execute("delete from shejixiadan where kehu_id="& CusID)
conn.execute("delete from kehu_jieri where kehu_id="& CusID)
conn.execute("delete from sjs_baobiao where xiangmu_id in (select id from shejixiadan where kehu_id="& CusID &")")
conn.execute("delete from richeng where kehu_id="& CusID)
conn.execute("delete from kehu where id="& CusID)
conn.execute("delete from chuzhu_jilu where kehu_id="& CusID)
End If
conn.execute("update shejixiadan s inner join save_money m on s.id=m.xiangmu_id set m.userid=s.userid where isnull(m.userid) or m.userid=''")
conn.execute("update shejixiadan set jixiang=0, jixiang_money=0 where isnull(jixiang)")
conn.execute("update CustomerCallInfo set times=now(),IsHangup=false where id=20")
response.write "更新完成"
%>