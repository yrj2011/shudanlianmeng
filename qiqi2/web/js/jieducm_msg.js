/* 
** ============================================
** 类名：CLASS_MSN_MESSAGE 
** 功能：提供类似MSN消息框  
**===========================================*/ 
var ttitle=document.title;
function CLASS_MSN_MESSAGE(id,width,height,caption,title,message,target,action){ 
this.id = id; 
this.title = title; 
this.caption= caption; 
this.message= message; 
this.target = target; 
this.action = action; 
this.width = width?width:200; 
this.height = height?height:60; 
this.timeout= 200; 
this.speed = 20; 
this.step = 1; 
this.right = screen.width -1; 
this.bottom = screen.height; 
this.left = this.right - this.width; 
this.top = this.bottom - this.height; 
this.timer = 0; 
this.pause = false; 
this.close = false; 
this.autoHide = true; 
} 

/**//* 
* 隐藏消息方法 
*/ 
CLASS_MSN_MESSAGE.prototype.hide = function(){ 
if(this.onunload()){ 

var offset = this.height>this.bottom-this.top?this.height:this.bottom-this.top; 
var me = this; 

if(this.timer>0){ 
window.clearInterval(me.timer); 
} 

var fun = function(){ 
if(me.pause==false||me.close){ 
var x = me.left; 
var y = 0; 
var width = me.width; 
var height = 0; 
if(me.offset>0){ 
height = me.offset; 
} 

y = me.bottom - height; 

if(y>=me.bottom){ 
window.clearInterval(me.timer); 
me.Pop.hide(); 
} else { 
me.offset = me.offset - me.step; 
} 
me.Pop.show(x,y,width,height); 
} 
} 

this.timer = window.setInterval(fun,this.speed) 
} 
document.title=ttitle;
} 

/**//* 
* 消息卸载事件，可以重写 
*/ 
CLASS_MSN_MESSAGE.prototype.onunload = function() { 
return true; 
} 
/**//* 
* 消息命令事件，要实现自己的连接，请重写它 
* 
*/ 
CLASS_MSN_MESSAGE.prototype.oncommand = function(){ 
//this.close = true; 
this.hide(); 
window.open("http://www.jieducm.com"); 

} 
/**//* 
* 消息显示方法 
*/ 
CLASS_MSN_MESSAGE.prototype.show = function(){ 

var oPopup = window.createPopup(); //IE5.5+ 

this.Pop = oPopup; 

var w = this.width; 
var h = this.height; 

var str = "<DIV style='BORDER-RIGHT: #455690 1px solid; BORDER-TOP: #a6b4cf 1px solid; Z-INDEX: 99999; LEFT: 0px; BORDER-LEFT: #a6b4cf 1px solid; WIDTH: 200px; BORDER-BOTTOM: #455690 1px solid; POSITION: absolute; TOP: 0px; HEIGHT: 120px; BACKGROUND-COLOR: #c9d3f3'>" 
str += "<TABLE style='BORDER-TOP: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid' cellSpacing=0 cellPadding=0 width='100%' bgColor=#cfdef4 border=0>" 
str += "<TR>" 
str += "<TD style='FONT-SIZE: 12px;COLOR: #0f2c8c' width=30 height=24></TD>" 
str += "<TD style='PADDING-LEFT: 4px; FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #1f336b; PADDING-TOP: 4px' vAlign=center width='100%'>" + this.caption + "</TD>" 
str += "<TD style='PADDING-RIGHT: 2px; PADDING-TOP: 2px' vAlign=center align=right width=19>" 
str += "<SPAN title=关闭 style='FONT-WEIGHT: bold; FONT-SIZE: 12px; CURSOR: hand; COLOR: red; MARGIN-RIGHT: 4px' id='btSysClose' >×</SPAN></TD>" 
str += "</TR>" 
str += "<TR>" 
str += "<TD style='PADDING-RIGHT: 1px;PADDING-BOTTOM: 1px' colSpan=3 height=30>" 
str += "<DIV style='BORDER-RIGHT: #b9c9ef 1px solid; PADDING-RIGHT: 8px; BORDER-TOP: #728eb8 1px solid; PADDING-LEFT: 8px; FONT-SIZE: 12px; PADDING-BOTTOM: 8px; BORDER-LEFT: #728eb8 1px solid; WIDTH: 100%; COLOR: #1f336b; PADDING-TOP: 8px; BORDER-BOTTOM: #b9c9ef 1px solid; HEIGHT: 100%'>" + this.title + "<BR><BR>" 
str += "</DIV>" 
str += "</TD>" 
str += "</TR>" 
str += "</TABLE>" 
str += "</DIV>" 
document.title="消息提示";
oPopup.document.body.innerHTML = str; 


this.offset = 0; 
var me = this; 

oPopup.document.body.onmouseover = function(){me.pause=true;} 
oPopup.document.body.onmouseout = function(){me.pause=false;} 

var fun = function(){ 
var x = me.left-15; 
var y = 0; 
var width = me.width; 
var height = me.height; 

if(me.offset>me.height){ 
height = me.height; 
} else { 
height = me.offset; 
} 

y = me.bottom - me.offset; 
if(y<=me.top){ 
me.timeout--; 
if(me.timeout==0){ 
window.clearInterval(me.timer); 
if(me.autoHide){ 
me.hide(); 
} 
} 
} else { 
me.offset = me.offset + me.step; 
} 
me.Pop.show(x,y,width,height); 

} 

this.timer = window.setInterval(fun,this.speed) 



var btClose = oPopup.document.getElementById("btSysClose"); 

btClose.onclick = function(){ 
me.close = true; 
me.hide(); 
} 

var btCommand = oPopup.document.getElementById("btCommand"); 
btCommand.onclick = function(){ 
me.oncommand(); 
} 
/*var ommand = oPopup.document.getElementById("ommand"); 
ommand.onclick = function(){ 
//this.close = true; 
me.hide(); 
window.open(ommand.href); 
} */
} 
/**//* 
** 设置速度方法 
**/ 
CLASS_MSN_MESSAGE.prototype.speed = function(s){ 
var t = 20; 
try { 
t = praseInt(s); 
} catch(e){} 
this.speed = t; 
} 
/**//* 
** 设置步长方法 
**/ 
CLASS_MSN_MESSAGE.prototype.step = function(s){ 
var t = 1; 
try { 
t = praseInt(s); 
} catch(e){} 
this.step = t; 
} 

CLASS_MSN_MESSAGE.prototype.rect = function(left,right,top,bottom){ 
try { 
this.left = left !=null?left:this.right-this.width; 
this.right = right !=null?right:this.left +this.width; 
this.bottom = bottom!=null?(bottom>screen.height?screen.height:bottom):screen.height; 
this.top = top !=null?top:this.bottom - this.height; 
} catch(e){} 
} 


function showmsg(mmsg)
{
var MSG1 = new CLASS_MSN_MESSAGE("aa",200,120,"消息提示：",mmsg); 
MSG1.rect(null,null,null,screen.height-50); 
MSG1.speed = 10; 
MSG1.step = 20; 
//alert(MSG1.top); 
MSG1.show();
}
//同时两个有闪烁，只能用层代替了，不过层不跨框架 
//var MSG2 = new CLASS_MSN_MESSAGE("aa",200,120,"短消息提示：","您有2封消息","ASP编程网欢迎您"); 
 //MSG2.rect(100,null,null,screen.height); 
 //MSG2.show(); 
//

//建立XMLHttpRequest对象
var xmlhttpmsg;
try{
    xmlhttpmsg= new ActiveXObject('Msxml2.xmlhttp');
}catch(e){
    try{
        xmlhttpmsg= new ActiveXObject('Microsoft.xmlhttp');
    }catch(e){
        try{
            xmlhttpmsg= new xmlhttpRequest();
        }catch(e){}
    }
}
var mmssgg="null";
function getmsg(u){
//document.getElementById("partdiv").innerHTML = "数据载入中。。。。。。。";
var username=u;
 //return 'sss';
var urlmsg="/inc/jieducm_msg.asp?"+"u="+username;
//alert(urlmsg);
    xmlhttpmsg.open("get",urlmsg,true);
    xmlhttpmsg.onreadystatechange = function(){

	 
	   
        if(xmlhttpmsg.readyState == 4)
        {   
		
            if(xmlhttpmsg.status == 200)
            {
				 
				
                if(xmlhttpmsg.responseText!=""){
                 
                   mmssgg = xmlhttpmsg.responseText;
				   
				   if(mmssgg!="null")
				 {
					showmsg(mmssgg);
				 }
             }
            }
            else{
                
				mmssgg = "null";
				
            }
        }
    }
    xmlhttpmsg.setRequestHeader("If-Modified-Since","0");
    xmlhttpmsg.send(null);




}

/*
var mmssgg='null';
function getmsg(u){
//document.getElementById("partdiv").innerHTML = "数据载入中。。。。。。。";
var username=u;

var urllq="umsg.asp?"+"u="+username;

	
		jQuery.ajax({
                        url:urllq,
                        Type:"post",
                        dataType:"text", 
                        beforeSend:function(){},
                        error:function(){
						mmssgg="null";
						//alert('bbb');
						},
                        timeout:190000,
                        success:function(data)
                        {
                               mmssgg=data;
							  
                        }
                    });
		if(mmssgg!="null")
				{
					showmsg(mmssgg);
				}
}
*/