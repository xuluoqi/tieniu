<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
%><HTML><HEAD><TITLE>客户项目专用面板</TITLE><META content=text/html; charset=gb2312 http-equiv=Content-Type><%
response.write "<script language='JavaScript'>"&vbcrlf
response.write "<!-- Begin"&vbcrlf
response.write "top.window.moveTo(-4,-4);"&vbcrlf
response.write "if (document.all) {"&vbcrlf
response.write "top.window.resizeTo(screen.availWidth+8,screen.availHeight+8);"&vbcrlf
response.write "}"&vbcrlf
response.write "else if (document.layers||document.getElementById) {"&vbcrlf
response.write "if (top.window.outerHeight<screen.availHeight||top.window.outerWidth&vbcrlf<screen.availWidth){"&vbcrlf
response.write "top.window.outerHeight = screen.availHeight;"&vbcrlf
response.write "top.window.outerWidth = screen.availWidth;"&vbcrlf
response.write "}"&vbcrlf
response.write "}"&vbcrlf
response.write "//  End -->"&vbcrlf
response.write "</script> "&vbcrlf
%><SCRIPT>var badwords=''</SCRIPT><STYLE type=text/css>
.navPoint {
	COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
</STYLE><SCRIPT>
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.all("frmTitle").style.display=""
}}
</SCRIPT><META content="MSHTML 5.00.3700.6699" name=GENERATOR><link href="../css/main.css" rel="stylesheet" type="text/css"></HEAD><BODY onresize=javascript:parent.carnoc.location.reload() scroll=no><CENTER><%
dim id,count11,rs
count11=0
id=trim(request("id"))
if id<>"" then
if isnumeric(id) then
count11=conn.execute("select count(*) from shejixiadan where isdelete=false and id="&id)(0)
end if
if count11<=0 then
set rs=conn.execute("select id from shejixiadan where isdelete=false and danhao='"&id&"'")
if not (rs.eof and rs.bof) then
response.write "<script language='javascript'>location.href='kehu_mianban.asp?id="&rs("id")&"'</script>"
On Error GoTo 0
Err.Raise 9999
end if
rs.close
set rs=conn.execute("select top 1 s.id from kehu k inner join shejixiadan s on k.id=s.kehu_id where k.isdelete=false and s.isdelete=false and k.number='"&id&"' order by s.id desc")
if not (rs.eof and rs.bof) then
response.write "<script language='javascript'>location.href='kehu_mianban.asp?id="&rs("id")&"'</script>"
On Error GoTo 0
Err.Raise 9999
end if
rs.close
set rs=nothing
response.write "<script>alert('对不起，没有找到该单号，请确认单号输入正确！');history.go(-1)</script>"
On Error GoTo 0
Err.Raise 9999
end if
end if
%><TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="100%"><TBODY><TR><TD height="100%" align=middle vAlign=center noWrap id=frmTitle name="frmTitle"><IFRAME   frameBorder=0 id=carnoc name=carnoc scrolling=no src="mianban_left.asp?id=<%= request("id") %>" 
      style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2"  ></IFRAME></TD><TD class=a2 style="WIDTH: 9pt"><TABLE border=0 cellPadding=0 cellSpacing=0 height="100%"><TBODY><TR><TD bgcolor="#5189FF" style="HEIGHT: 100%" onclick=switchSysBar()><FONT 
            style="COLOR: #ffffff; CURSOR: default; FONT-SIZE: 9pt"><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><SPAN 
            class=navPoint id=switchPoint 
            title=关闭/打开左栏>3</SPAN><BR><BR><BR><BR><BR><BR><BR><BR>屏幕切换 
        </FONT></TD></TR></TBODY></TABLE></TD><TD valign="top" bgcolor="#FFFFFF" style="WIDTH: 100%"><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0"><tr><td height="100%"><IFRAME frameBorder=0 height=100% marginHeight=1 marginWidth=1 src="<%
if request("id")<>"" then response.write "admin/wancheng_print.asp?id="&request("id")
%>" width="100%" BORDERCOLOR="#000000" name="main2" scrolling=auto"  align="center"  id="llz"></IFRAME></td></tr></table></TD></TR></TBODY></TABLE></BODY></HTML>