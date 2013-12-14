
function setTab(m,n){
 var tli=document.getElementById("b"+m).getElementsByTagName("li");
 var mli=document.getElementById("a"+m).getElementsByTagName("dl");
 for(i=0;i<tli.length;i++){
  tli[i].className=i==n?"hover":"";
  mli[i].style.display=i==n?"block":"none";
 }
}

