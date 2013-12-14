<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
conn.execute("delete from shejixiadan where jixiang not in (select id from jixiang )")
conn.execute("delete from fujia where xiangmu_id not in (select id from shejixiadan )")
conn.execute("delete from fujia2 where xiangmu_id not in (select id from shejixiadan )")
conn.execute("delete from sjs_baobiao where xiangmu_id not in (select id from shejixiadan )")
conn.execute("delete from goumai where xiangmu_id not in (select id from shejixiadan )")
conn.execute("delete from xiadan where xiangmu_id not in (select id from shejixiadan )")
%>