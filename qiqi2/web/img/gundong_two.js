document.writeln("	<div class=\"rollBox\">");
document.writeln("		 <div class=\"LeftBotton\" onmousedown=\"ISL_GoUp()\" onmouseup=\"ISL_StopUp()\" onmouseout=\"ISL_StopUp()\"><\/div>");
document.writeln("		 <div class=\"Cont\" id=\"ISL_Cont\">");
document.writeln("		  <div class=\"ScrCont\">");
document.writeln("		   <div id=\"List1\">");
document.writeln("		   ");
document.writeln("			<!-- 图片列表 begin -->");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.webgle.com\" target=\"_blank\"><img src=\"images\/webgle.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.webglr.com\" target=\"_blank\"><img src=\"images\/webglr.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.webgld.com\" target=\"_blank\"><img src=\"images\/webgld.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.webgll.com\" target=\"_blank\"><img src=\"images\/webgll.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.csd198.com\" target=\"_blank\"><img src=\"images\/csd198.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/yaya.webgle.com\" target=\"_blank\"><img src=\"images\/yaya.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.taobao.com\" target=\"_blank\"><img src=\"images\/taobao.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.alipay.com\" target=\"_blank\"><img src=\"images\/alipay.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.alibaba.com\" target=\"_blank\"><img src=\"images\/alibaba.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.alimama.com\" target=\"_blank\"><img src=\"images\/alimama.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.paipai.com\" target=\"_blank\"><img src=\"images\/paipai.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.tenpay.com\" target=\"_blank\"><img src=\"images\/tenpay.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.youa.com\" target=\"_blank\"><img src=\"images\/youa.gif\" \/><\/a><\/div>       ");
document.writeln("				<div class=\"pic\"><a href=\"http:\/\/www.baifubao.com\" target=\"_blank\"><img src=\"images\/baifubao.gif\" \/><\/a><\/div>       ");
document.writeln("			<!-- 图片列表 end -->");
document.writeln("			");
document.writeln("		   <\/div>");
document.writeln("		   <div id=\"List2\"><\/div>");
document.writeln("		  <\/div>");
document.writeln("		 <\/div>");
document.writeln("		 <div class=\"RightBotton\" onmousedown=\"ISL_GoDown()\" onmouseup=\"ISL_StopDown()\" onmouseout=\"ISL_StopDown()\"><\/div>");
document.writeln("		<\/div>");
document.writeln("")
//图片滚动列表 mengjia 070816
var Speed = 10; //速度(毫秒)
var Space = 5; //每次移动(px)
var PageWidth = 948; //翻页宽度
var fill = 0; //整体移位
var MoveLock = false;
var MoveTimeObj;
var Comp = 0;
var AutoPlayObj = null;
GetObj("List2").innerHTML = GetObj("List1").innerHTML;
GetObj('ISL_Cont').scrollLeft = fill;
GetObj("ISL_Cont").onmouseover = function(){clearInterval(AutoPlayObj);}
GetObj("ISL_Cont").onmouseout = function(){AutoPlay();}
AutoPlay();
function GetObj(objName){if(document.getElementById){return eval('document.getElementById("'+objName+'")')}else{return eval('document.all.'+objName)}}
function AutoPlay(){ //自动滚动
 clearInterval(AutoPlayObj);
 AutoPlayObj = setInterval('ISL_GoDown();ISL_StopDown();',10000); //间隔时间
}
function ISL_GoUp(){ //上翻开始
 if(MoveLock) return;
 clearInterval(AutoPlayObj);
 MoveLock = true;
 MoveTimeObj = setInterval('ISL_ScrUp();',Speed);
}
function ISL_StopUp(){ //上翻停止
 clearInterval(MoveTimeObj);
 if(GetObj('ISL_Cont').scrollLeft % PageWidth - fill != 0){
  Comp = fill - (GetObj('ISL_Cont').scrollLeft % PageWidth);
  CompScr();
 }else{
  MoveLock = false;
 }
 AutoPlay();
}
function ISL_ScrUp(){ //上翻动作
 if(GetObj('ISL_Cont').scrollLeft <= 0){GetObj('ISL_Cont').scrollLeft = GetObj('ISL_Cont').scrollLeft + GetObj('List1').offsetWidth}
 GetObj('ISL_Cont').scrollLeft -= Space ;
}
function ISL_GoDown(){ //下翻
 clearInterval(MoveTimeObj);
 if(MoveLock) return;
 clearInterval(AutoPlayObj);
 MoveLock = true;
 ISL_ScrDown();
 MoveTimeObj = setInterval('ISL_ScrDown()',Speed);
}
function ISL_StopDown(){ //下翻停止
 clearInterval(MoveTimeObj);
 if(GetObj('ISL_Cont').scrollLeft % PageWidth - fill != 0 ){
  Comp = PageWidth - GetObj('ISL_Cont').scrollLeft % PageWidth + fill;
  CompScr();
 }else{
  MoveLock = false;
 }
 AutoPlay();
}
function ISL_ScrDown(){ //下翻动作
 if(GetObj('ISL_Cont').scrollLeft >= GetObj('List1').scrollWidth){GetObj('ISL_Cont').scrollLeft = GetObj('ISL_Cont').scrollLeft - GetObj('List1').scrollWidth;}
 GetObj('ISL_Cont').scrollLeft += Space ;
}
function CompScr(){
 var num;
 if(Comp == 0){MoveLock = false;return;}
 if(Comp < 0){ //上翻
  if(Comp < -Space){
   Comp += Space;
   num = Space;
  }else{
   num = -Comp;
   Comp = 0;
  }
  GetObj('ISL_Cont').scrollLeft -= num;
  setTimeout('CompScr()',Speed);
 }else{ //下翻
  if(Comp > Space){
   Comp -= Space;
   num = Space;
  }else{
   num = Comp;
   Comp = 0;
  }
  GetObj('ISL_Cont').scrollLeft += num;
  setTimeout('CompScr()',Speed);
 }
}
