<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from yuangong where username='"&session("userid")&"' and password='"&session("password")&"'",conn,1,1
if rs.eof and rs.bof then
response.write "<SCRIPT language=JavaScript>alert('对不起，你没有权限进入该页面!');"
response.write"this.location.href='index.asp';</SCRIPT>"
On Error GoTo 0
Err.Raise 9999
end if
rs.close
set rs=nothing
if  session("level")="" then
response.write "<SCRIPT language=JavaScript>alert('对不起，你没有权限进入该页面!');"
response.write"this.location.href='index.asp';</SCRIPT>"
On Error GoTo 0
Err.Raise 9999
end if
session.Timeout=60
%><HTML><HEAD><TITLE>管理员操作专区</TITLE><META content=text/html; charset=gb2312 charset=gb2312 http-equiv=Content-Type><script type="text/JavaScript"><!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function getURL(id){
	parent.carnoc.location.href='admin/left.asp?show='+id;
}
//--></script><style type="text/css"><!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
--></style></head><body onLoad="MM_preloadImages('Image/ERP2_02.jpg','Image/ERP2_03.jpg','Image/ERP2_04.jpg','Image/ERP2_05.jpg','Image/ERP2_06.jpg','Image/ERP2_07.jpg','Image/ERP2_08.jpg','Image/ERP2_09.jpg')"><table border="0" cellpadding="0" cellspacing="0"><tr><td><img src="Image/ERP201.jpg" width="315" height="104" /></td><td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','Image/ERP2_02.jpg',1)"><img src="Image/ERP202.jpg" name="Image3" width="89" height="104" border="0" id="Image3" /></a></td><td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image4','','Image/ERP2_03.jpg',1)"><img src="Image/ERP203.jpg" name="Image4" width="85" height="104" border="0" id="Image4" /></a></td><td><a href="#" onMouseOver="MM_swapImage('Image5','','Image/ERP2_04.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="Image/ERP204.jpg" name="Image5" width="87" height="104" border="0" id="Image5" /></a></td><td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image6','','Image/ERP2_05.jpg',1)"><img src="Image/ERP205.jpg" name="Image6" width="88" height="104" border="0" id="Image6" /></a></td><td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image7','','Image/ERP2_06.jpg',1)"><img src="Image/ERP206.jpg" name="Image7" width="87" height="104" border="0" id="Image7" /></a></td><td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','Image/ERP2_07.jpg',1)"><img src="Image/ERP207.jpg" name="Image8" width="85" height="104" border="0" id="Image8" /></a></td><td><a href="#" target="main" onMouseOver="MM_swapImage('Image9','','Image/ERP2_08.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="Image/ERP208.jpg" name="Image9" width="92" height="104" border="0" id="Image9" /></a></td><td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image10','','Image/ERP2_09.jpg',1)"><img src="Image/ERP209.jpg" name="Image10" width="77" height="104" border="0" id="Image10" /></a></td></tr></table></body></html>