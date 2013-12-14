<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
if session("level")="" then
response.Write "<script>alert('对不起，你没有权限进入该页面！');history.go(-1)</script>"
On Error GoTo 0
Err.Raise 9999
end if
%><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>站内公告信息录入</title><link href="../admin/zxcss.css" rel="stylesheet" type="text/css"><link href="admin/zxcss.css" rel="stylesheet" type="text/css"><script src="js/AC_ActiveX.js" type="text/javascript"></script><script src="js/AC_RunActiveContent.js" type="text/javascript"></script><body marginheight=0 marginwidth=0 leftmargin=0 ><script LANGUAGE="JavaScript">
function check()
{
document.Form1.content.value=document.Form1.content_html.value;
if (document.Form1.title.value=="")
{
alert("请输入公告标题！")
document.Form1.title.focus()
document.Form1.title.select()
return
}
if (document.Form1.topeple.value=="")
{
alert("请选择留言对象！")
document.Form1.topeple.focus()
document.Form1.topeple.select()
return
}
if (document.Form1.content_html.value=="")
{
alert("请输入文章内容！")
return
}
if (document.Form1.pic.value!=""){
   if(( document.Form1.pic.value.indexOf(".gif") == -1) && (document.Form1.pic.value.indexOf(".jpg") == -1) && (document.Form1.pic.value.indexOf(".JPG") == -1) && (document.Form1.pic.value.indexOf(".GIF") == -1)) 
        {
        alert("请选择gif或jpg的图象文件！");
		document.Form1.pic.focus();
        return (false);
        }
}
document.Form1.submit()
}
</SCRIPT><CENTER>　
  <TABLE width="100%" border="0" align="center" cellspacing="1" bordercolor="#111111" class="border" style="border-collapse: collapse"><TR style="background-image: url('../../Images/topbg1.gif')"><TD width="990" height="25" align="center" class="F12"><strong>
  员工留言版</strong></TD></TR></TABLE><TABLE width=100% border="0" align="center" cellpadding="0" cellspacing="0" height=403><TR><TD width="1009" height="403"><FORM action=AddInfo.asp method="POST" name="Form1" style="margin:0px " enctype="multipart/form-data"><TABLE width="100%" border="0" cellpadding="2" cellspacing="1" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" class="border" style="border-collapse: collapse"><TR class="tdbg"><TD align="right" width="106" height="25"><div align="center">标题：</div></TD><TD width="885" height="25">
              &nbsp;
              <input name=title type=text class="smallInput" id="title" size="64"></TD></TR><TR class="tdbg"><TD height="25" align="right"><div align="center">留言对象：</div></TD><TD width="885" height="25" bgcolor="#E1F4EE"><font color="#FFFFFF">
&nbsp;                <select name="topeple" id="topeple"><%
set rs=server.CreateObject("adodb.recordset")
sql="select * from yuangong where peplename<>'"&session("username")&"'"
rs.open sql,connstr,1,1
%><option value="">请选择</option><%
while not rs.eof
%><option value="<%= rs("peplename") %>"><%= rs("peplename") %></option><%
rs.movenext
wend
rs.close
set rs=nothing
%></select></font></TD></TR><TR class="tdbg"><TD height="25" align="right"><div align="center">日期:</div></TD><TD height="25"><font color="#FFFFFF">
                &nbsp;
                <input name=idate type=text class="smallInput" id="idate" size="30" value="<%=  now %>">
              </font></TD></TR><tr class="tdbg"><TD height="21" align="right"><div align="center"><TEXTAREA STYLE="display:none" NAME="content"></TEXTAREA></div></TD><TD height="21" align="right"><object id="content_html" style="LEFT: 0px; TOP: 0px" data="inc2/edit.htm" width=100% height=273 type=text/x-scriptlet  viewastext><embed src="INC2/edit.htm" width="100%" height="273"></embed></object></TD></tr><tr class="tdbg"><TD height="28" align="right"><div align="center">上传图片:              </div></TD><TD height="28" align="right"><div align="left"><input name="pic" type="file" class="bgcolor" id="pic" style="width:210" value=""><input type="hidden" name="act" value="uploadfile"><input type="hidden" name="frompeple" value="<%= session("username") %>">
</div></TD></tr><TR class="tdbg"><TD height="20" align="center" colspan="2"><input name="button" type="button" class="bgcolor"  onclick=check() value="添 加">
          &nbsp;&nbsp;&nbsp;&nbsp;
                <input name="button" type="button" class="bgcolor" onclick=javascript:history.go(-1) value="返 回"></TD></TR></TABLE><font color="#FFFFFF"></font></FORM></TD></TR></TABLE>