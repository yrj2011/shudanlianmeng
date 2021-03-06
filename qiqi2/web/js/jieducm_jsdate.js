
 <!--
var cal;
var isFocus=false; //是否为焦点
//以上为  寒羽枫 2006-06-25 添加的变量

//选择日期 → 由 寒羽枫 2006-06-25 添加
function SelectDate(obj,strFormat)
{   
    var date = new Date();
    var by = date.getFullYear()-0;  //最小值 → 50 年前
    var ey = date.getFullYear()+20;  //最大值 → 50 年后
    //cal = new Calendar(by, ey,1,strFormat);    //初始化英文版，0 为中文版
    cal = (cal==null) ? new Calendar(by, ey, 1) : cal;    //不用每次都初始化 2006-12-03 修正      
 cal.dateFormatStyle = strFormat;
    cal.show(obj);
}
/**//**//**//**
 * 返回日期
 * @param d the delimiter
 * @param p the pattern of your date
 2006-06-25 由 寒羽枫 修改为根据用户指定的 style 来确定；
 */
//String.prototype.toDate = function(x, p) {
String.prototype.toDate = function(style) {
/**//**//**//*
  if(x == null) x = "-";
  if(p == null) p = "ymd";
  var a = this.split(x);
  var y = parseInt(a[p.indexOf("y")]);
  //remember to change this next century ;)
  if(y.toString().length <= 2) y += 2000;
  if(isNaN(y)) y = new Date().getFullYear();
  var m = parseInt(a[p.indexOf("m")]) - 1;
  var d = parseInt(a[p.indexOf("d")]);
  if(isNaN(d)) d = 1;
  return new Date(y, m, d);
  */
  var y = this.substring(style.indexOf('y'),style.lastIndexOf('y')+1);//年
  var m = this.substring(style.indexOf('M'),style.lastIndexOf('M')+1);//月
  var d = this.substring(style.indexOf('d'),style.lastIndexOf('d')+1);//日


  var h = this.substring(style.indexOf('h'),style.lastIndexOf('h')+1);
  var mi = this.substring(style.indexOf('m'),style.lastIndexOf('m')+1);
  var s = this.substring(style.indexOf('s'),style.lastIndexOf('s')+1);
  
  if(isNaN(y)) y = new Date().getFullYear();
  if(isNaN(m)) m = new Date().getMonth();
  if(isNaN(d)) d = new Date().getDate();
  if(isNaN(h)) h = new Date().getHours();
  if(isNaN(mi)) mi = new Date().getMinutes();
  if(isNaN(s)) s = new Date().getSeconds();
  var dt ;
  eval ("dt = new Date('"+ y+"', '"+(m-1)+"','"+ d +"','"+h+"','"+mi+"','"+s+"')");
//eval ("dt = new Date('"+ y+"', '"+(m-1)+"','"+ d +"')");
  //eval("dt = new Date(2008,1,2)");
  return dt;
}

/**//**//**//**
 * 格式化日期
 * @param   d the delimiter
 * @param   p the pattern of your date
 * @author  meizz
 */
Date.prototype.format = function(style) {
  var o = {
    "M+" : this.getMonth() + 1, //month
    "d+" : this.getDate(),      //day
    "h+" : this.getHours(),     //hour
    "m+" : this.getMinutes(),   //minute
    "s+" : this.getSeconds(),   //second
    "w+" : "天一二三四五六".charAt(this.getDay()),   //week
    "q+" : Math.floor((this.getMonth() + 3) / 3),  //quarter
    "S"  : this.getMilliseconds() //millisecond
  }
  if(/(y+)/.test(style)) {
    style = style.replace(RegExp.$1,
    (this.getFullYear() + "").substr(4 - RegExp.$1.length));
  }
  for(var k in o){
    if(new RegExp("("+ k +")").test(style)){
      style = style.replace(RegExp.$1,
        RegExp.$1.length == 1 ? o[k] :
        ("00" + o[k]).substr(("" + o[k]).length));
    }
  }
  return style;
};

/**//**//**//**
 * 日历类
 * @param   beginYear 1990
 * @param   endYear   2010
 * @param   lang      0(中文)|1(英语) 可自由扩充
 * @param   dateFormatStyle  "yyyy-MM-dd";
 * @version 2006-04-01
 * @author  KimSoft (jinqinghua [at] gmail.com)
 * @update
 */
function Calendar(beginYear, endYear, lang, dateFormatStyle) {
  this.beginYear = 1990;
  this.endYear = 2010;
  this.lang = 0;            //0(中文) | 1(英文)
  
 //改变默认日期时间格式 mazq 修改
 // this.dateFormatStyle = "yyyy-MM-dd";
 this.dateFormatStyle="yyyy-MM-dd hh:mm:ss";

 //this.calendarTime=new minute(this.id, this);
 //以上 mazq 修改

  if (beginYear != null && endYear != null){
    this.beginYear = beginYear;
    this.endYear = endYear;
  }
  if (lang != null){
    this.lang = lang
  }

  if (dateFormatStyle != null){
    this.dateFormatStyle = dateFormatStyle
  }

  this.dateControl = null;
  this.panel = this.getElementById("calendarPanel");
  this.container = this.getElementById("ContainerPanel");
  this.form  = null;

  this.date = new Date();
  this.year = this.date.getFullYear();
  this.month = this.date.getMonth();
  /*以下代码mazq 添加 2007-04-02*/
  this.hours= this.date.getHours();
  this.minutes = this.date.getMinutes();
  this.seconds = this.date.getSeconds();

  this.timer = null;
  this.fObj = null;
  /*以上代码mazq 添加*/

  this.colors = {
  "cur_word"      : "#FFFFFF",  //当日日期文字颜色
  "cur_bg"        : "#00FF00",  //当日日期单元格背影色
  "sel_bg"        : "#FFCCCC",  //已被选择的日期单元格背影色 2006-12-03 寒羽枫添加
  "sun_word"      : "#FF0000",  //星期天文字颜色
  "sat_word"      : "#0000FF",  //星期六文字颜色
  "td_word_light" : "#333333",  //单元格文字颜色
  "td_word_dark"  : "#CCCCCC",  //单元格文字暗色
  "td_bg_out"     : "#EFEFEF",  //单元格背影色
  "td_bg_over"    : "#FFCC00",  //单元格背影色
  "tr_word"       : "#FFFFFF",  //日历头文字颜色
  "tr_bg"         : "#666666",  //日历头背影色
  "input_border"  : "#CCCCCC",  //input控件的边框颜色
  "input_bg"      : "#EFEFEF"   //input控件的背影色
  }

  this.draw();
  this.bindYear();
  this.bindMonth();
  this.changeSelect();
  this.bindData();
}

/**//**//**//**
 * 日历类属性（语言包，可自由扩展）
 */
Calendar.language = {
  "year"   : [[""], [""]],
  "months" : [["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"],
        ["一月","二月","三月","四月","五月","六月","七月","八月","九月","十月","十一月","十二月"]
         ],
  "weeks"  : [["SUN","MON","TUR","WED","THU","FRI","SAT"],
        ["日","一","二","三","四","五","六"]
         ],
  "clear"  : [["CLS"], ["清空"]],
  "today"  : [["TODAY"], ["今天"]],
  "close"  : [["CLOSE"], ["关闭"]]
}

Calendar.prototype.draw = function() {
  calendar = this;

  var mvAry = [];
  //mvAry[mvAry.length]  = '  <form name="calendarForm" style="margin: 0px;">'; //因 <form> 不能嵌套， 2006-12-01 由寒羽枫改用 Div
  mvAry[mvAry.length]  = '  <div name="calendarForm" style="margin: 0px;z-index:20000">';
  mvAry[mvAry.length]  = '    <table width="100%" border="0" cellpadding="0" cellspacing="1">';
  mvAry[mvAry.length]  = '      <tr>';
  mvAry[mvAry.length]  = '        <th align="left" width="1%"><input style="border: 1px solid ' + calendar.colors["input_border"] + ';background-color:' + calendar.colors["input_bg"] + ';width:16px;height:20px;" name="prevMonth" type="button" id="prevMonth" value="&lt;" /></th>';
  mvAry[mvAry.length]  = '        <th align="center" width="98%" nowrap="nowrap"><select name="calendarYear" id="calendarYear" style="font-size:10px;"></select><select name="calendarMonth" id="calendarMonth" style="font-size:10px;"></select></th>';
  mvAry[mvAry.length]  = '        <th align="right" width="1%"><input style="border: 1px solid ' + calendar.colors["input_border"] + ';background-color:' + calendar.colors["input_bg"] + ';width:16px;height:20px;" name="nextMonth" type="button" id="nextMonth" value="&gt;" /></th>';
  mvAry[mvAry.length]  = '      </tr>';
  mvAry[mvAry.length]  = '    </table>';
  mvAry[mvAry.length]  = '    <table id="calendarTable" width="100%" style="border:0px solid #CCCCCC;background-color:#FFFFFF" border="0" cellpadding="3" cellspacing="1">';
  mvAry[mvAry.length]  = '      <tr>';
  for(var i = 0; i < 7; i++) {
    mvAry[mvAry.length]  = '      <th style="font-weight:normal;font-size:10px;background-color:' + calendar.colors["tr_bg"] + ';color:' + calendar.colors["tr_word"] + ';">' + Calendar.language["weeks"][this.lang][i] + '</th>';
  }
  mvAry[mvAry.length]  = '      </tr>';
  for(var i = 0; i < 6;i++){
    mvAry[mvAry.length]  = '    <tr align="center">';
    for(var j = 0; j < 7; j++) {
      if (j == 0){
        mvAry[mvAry.length]  = '  <td style="cursor:default;font-size:10px;width:14px;height:18px;color:' + calendar.colors["sun_word"] + ';"></td>';
      } else if(j == 6) {
        mvAry[mvAry.length]  = '  <td style="cursor:default;font-size:10px;width:14px;height:18px;color:' + calendar.colors["sat_word"] + ';"></td>';
      } else {
        mvAry[mvAry.length]  = '  <td style="cursor:default;font-size:10px;width:14px;height:18px;"></td>';
      }
    }
    mvAry[mvAry.length]  = '    </tr>';
  }
  mvAry[mvAry.length]  = '      <tr style="background-color:' + calendar.colors["input_bg"] + ';">';
  mvAry[mvAry.length]  = '        <th colspan="2"><input name="calendarClear" type="button" id="calendarClear" value="' + Calendar.language["clear"][this.lang] + '" style="border: 1px solid ' + calendar.colors["input_border"] + ';background-color:' + calendar.colors["input_bg"] + ';width:100%;height:20px;font-size:10px;"/></th>';
  mvAry[mvAry.length]  = '        <th colspan="3"><input name="calendarToday" type="button" id="calendarToday" value="' + Calendar.language["today"][this.lang] + '" style="border: 1px solid ' + calendar.colors["input_border"] + ';background-color:' + calendar.colors["input_bg"] + ';width:100%;height:20px;font-size:10px;"/></th>';
  mvAry[mvAry.length]  = '        <th colspan="2"><input name="calendarClose" type="button" id="calendarClose" value="' + Calendar.language["close"][this.lang] + '" style="border: 1px solid ' + calendar.colors["input_border"] + ';background-color:' + calendar.colors["input_bg"] + ';width:100%;height:20px;font-size:10px;"/></th>';
  mvAry[mvAry.length]  = '      </tr>';
  
  /* 以下mazq添加 */  
  mvAry[mvAry.length]  = '      <tr>';
  mvAry[mvAry.length]  = '      <td colspan="7" align="center">';
  var sMinute_Common = "style='width: 18px;height: 14px;border: 0px solid black;font-family: Tahoma;font-size: 9px;text-align: right;ime-mode:disabled' maxlength='2' id='txtcalendarTime' name='calendarTime' onfocus='calendar.setFocusObj(this)' onblur='calendar.setTime(this)' onkeyup='calendar.prevent(this)' onkeypress='if (!/[0-9]/.test(String.fromCharCode(event.keyCode)))event.keyCode=0' onpaste='return false' ondragenter='return false'";
  var sButton_Common = "style='width: 16px;height: 8px;font-family: Webdings;font-size: 7px;line-height: 2px;padding-left: 2px;cursor: default' onfocus='this.blur()' onmouseup='calendar.controlTime()' disabled"
  var str = "";
  str += "<table name=\"c_time\" id =\"c_time\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">"
  str += " <tr>"
  str += "  <td>"
  str += "   <span style='vertical-align:middle;font-family: Arial;font-size: 9pt;'>时间</span>"
  str += "  </td>"
  str += "  <td>"
  str += "   <div style=\"border-left: 2px inset #D4D0C8;border-top: 2px inset #D4D0C8;border-right: 2px inset #FFFFFF;border-bottom: 2px inset #FFFFFF;width: 100px;height: 19px;background-color: #FFFFFF;overflow: hidden;text-align: right;font-family: Tahoma;font-size: 10px;\">"
  str += "    <input radix=\"24\" name=\"c_hours\"  id= \"c_hours\" value=\"\""+sMinute_Common+">:"
  str += "    <input radix=\"60\" name=\"c_minutes\" id= \"c_minutes\" value=\"\""+sMinute_Common+">:"
  str += "    <input radix=\"60\" name=\"c_seconds\" id= \"c_seconds\" value=\"\""+sMinute_Common+">"
  str += "   </div>"
  str += "  </td>"
  str += "  <td>"
  str += "  <table border=\"0\" cellspacing=\"2\" cellpadding=\"0\">"
  str += "   <tr><td><button id=\"c_up\" "+sButton_Common+">5</button></td></tr>"
  str += "   <tr><td><button id=\"c_down\" "+sButton_Common+">6</button></td></tr>"
  str += "  </table>"
  str += "  </td>"
  str += " </tr>"
  str += "</table>"
  mvAry[mvAry.length]  = str;
  mvAry[mvAry.length]  = '      </td>';
  mvAry[mvAry.length]  = '    </tr>';
   /* 以上mazq添加 */
  mvAry[mvAry.length]  = '    </table>';
  //mvAry[mvAry.length]  = '  </from>';
  mvAry[mvAry.length]  = '  </div>';

  this.panel.innerHTML = mvAry.join("");
  
  
  /**//******** 以下代码由寒羽枫 2006-12-01 添加 **********/
  var obj = this.getElementById("prevMonth");
  obj.onclick = function () {calendar.goPrevMonth(calendar);}
  obj.onblur = function () {calendar.onblur();}
  this.prevMonth= obj;
  
  obj = this.getElementById("nextMonth");
  obj.onclick = function () {calendar.goNextMonth(calendar);}
  obj.onblur = function () {calendar.onblur();}
  this.nextMonth= obj;
  
  

  obj = this.getElementById("calendarClear");
  obj.onclick = function () {calendar.dateControl.value = "";calendar.hide();}
  this.calendarClear = obj;
  
  obj = this.getElementById("calendarClose");
  obj.onclick = function () {calendar.hide();}
  this.calendarClose = obj;
  
  obj = this.getElementById("calendarYear");
  obj.onchange = function () {calendar.update(calendar);}
  obj.onblur = function () {calendar.onblur();}
  this.calendarYear = obj;
  
  obj = this.getElementById("calendarMonth");
  with(obj)
  {
    onchange = function () {calendar.update(calendar);}
    onblur = function () {calendar.onblur();}
  }this.calendarMonth = obj;
 
  obj = this.getElementById("calendarToday");
  obj.onclick = function () {
    var today = new Date();
    calendar.date = today;
    calendar.year = today.getFullYear();
    calendar.month = today.getMonth();
    
 //以下 mazq 2007-4-4 添加
    calendar.hours =today.getHours();
    calendar.minutes = today.getMinutes();
    calendar.seconds = today.getSeconds();
    //以上 mazq 2007-4-4 添加
    calendar.changeSelect();
    calendar.bindData();
    calendar.dateControl.value = today.format(calendar.dateFormatStyle);
    calendar.hide();
  }
  this.calendarToday = obj;
  /**//******** 以上代码由寒羽枫 2006-12-01 添加 **********/
  

  /**//*
  //this.form = document.forms["calendarForm"];   
  this.form.prevMonth.onclick = function () {calendar.goPrevMonth(this);}
  this.form.nextMonth.onclick = function () {calendar.goNextMonth(this);}
  
  this.form.prevMonth.onblur = function () {calendar.onblur();}
  this.form.nextMonth.onblur = function () {calendar.onblur();}

  this.form.calendarClear.onclick = function () {calendar.dateControl.value = "";calendar.hide();}
  this.form.calendarClose.onclick = function () {calendar.hide();}
  this.form.calendarYear.onchange = function () {calendar.update(this);}
  this.form.calendarMonth.onchange = function () {calendar.update(this);}
  
  this.form.calendarYear.onblur = function () {calendar.onblur();}
  this.form.calendarMonth.onblur = function () {calendar.onblur();}
  
  this.form.calendarToday.onclick = function () {
    var today = new Date();
    calendar.date = today;
    calendar.year = today.getFullYear();
    calendar.month = today.getMonth();
    calendar.changeSelect();
    calendar.bindData();
    calendar.dateControl.value = today.format(calendar.dateFormatStyle);
    calendar.hide();
  }
*/
}

//年份下拉框绑定数据
Calendar.prototype.bindYear = function() {
  //var cy = this.form.calendarYear;
  var cy = this.calendarYear;//2006-12-01 由寒羽枫修改
  cy.length = 0;
  for (var i = this.beginYear; i <= this.endYear; i++){
    cy.options[cy.length] = new Option(i + Calendar.language["year"][this.lang], i);
  }
}

//月份下拉框绑定数据
Calendar.prototype.bindMonth = function() {
  //var cm = this.form.calendarMonth;
  var cm = this.calendarMonth;//2006-12-01 由寒羽枫修改
  cm.length = 0;
  for (var i = 0; i < 12; i++){
    cm.options[cm.length] = new Option(Calendar.language["months"][this.lang][i], i);
  }
}

//向前一月
Calendar.prototype.goPrevMonth = function(e){
  if (this.year == this.beginYear && this.month == 0){return;}
  this.month--;
  if (this.month == -1) {
    this.year--;
    this.month = 11;
  }
  this.date = new Date(this.year, this.month, 1);
  this.changeSelect();
  this.bindData();
}

//向后一月
Calendar.prototype.goNextMonth = function(e){
  if (this.year == this.endYear && this.month == 11){return;}
  this.month++;
  if (this.month == 12) {
    this.year++;
    this.month = 0;
  }
  this.date = new Date(this.year, this.month, 1,this.hours,this.minutes,this.seconds);
  this.changeSelect();
  this.bindData();
}

//改变SELECT选中状态
Calendar.prototype.changeSelect = function() {
  //var cy = this.form.calendarYear;
  //var cm = this.form.calendarMonth;
  var cy = this.calendarYear;//2006-12-01 由寒羽枫修改
  var cm = this.calendarMonth;
  for (var i= 0; i < cy.length; i++){
    if (cy.options[i].value == this.date.getFullYear()){
      cy[i].selected = true;
      break;
    }
  }
  for (var i= 0; i < cm.length; i++){
    if (cm.options[i].value == this.date.getMonth()){
      cm[i].selected = true;
      break;
    }
  }
}

//更新年、月
Calendar.prototype.update = function (e){
  //this.year  = e.form.calendarYear.options[e.form.calendarYear.selectedIndex].value;
  //this.month = e.form.calendarMonth.options[e.form.calendarMonth.selectedIndex].value;
  this.year  = e.calendarYear.options[e.calendarYear.selectedIndex].value;//2006-12-01 由寒羽枫修改
  this.month = e.calendarMonth.options[e.calendarMonth.selectedIndex].value;
  this.date = new Date(this.year, this.month, 1,this.hours,this.minutes,this.seconds);
  this.changeSelect();
  this.bindData();
}

//绑定数据到月视图
Calendar.prototype.bindData = function () {
  var calendar = this;
  var dateArray = this.getMonthViewArray(this.date.getYear(), this.date.getMonth());
  var tds = this.getElementById("calendarTable").getElementsByTagName("td");
  for(var i = 0; i < tds.length; i++) {
  //tds[i].style.color = calendar.colors["td_word_light"];
  tds[i].style.backgroundColor = calendar.colors["td_bg_out"];
    tds[i].onclick = function () {return;}
    tds[i].onmouseover = function () {return;}
    tds[i].onmouseout = function () {return;}
    if (i > dateArray.length - 1) break;
    tds[i].innerHTML = dateArray[i];
    if (dateArray[i] != "&nbsp;"){
      tds[i].onclick = function () {
        if (calendar.dateControl != null){
         // calendar.dateControl.value = new Date(calendar.date.getFullYear(),                                           calendar.date.getMonth(),this.innerHTML).format(calendar.dateFormatStyle);

  //以下 mazq 2007-4-4 添加和修改
 this.hours = c_hours.value;
 this.minutes = c_minutes.value;
 this.seconds = c_seconds.value;   
   this.date =  new Date(calendar.date.getFullYear(),            calendar.date.getMonth(),this.innerHTML,this.hours,this.minutes,this.seconds);       
calendar.dateControl.value = this.date.format(calendar.dateFormatStyle);
  //以上 mazq 2007-4-4 添加和修改

        }
        calendar.hide();
      }
      tds[i].onmouseover = function () {
        this.style.backgroundColor = calendar.colors["td_bg_over"];
      }
      tds[i].onmouseout = function () {
        this.style.backgroundColor = calendar.colors["td_bg_out"];
      }
      if (new Date().format(calendar.dateFormatStyle) ==
          new Date(calendar.date.getFullYear(),
                   calendar.date.getMonth(),
                   dateArray[i]).format(calendar.dateFormatStyle)) {
        //tds[i].style.color = calendar.colors["cur_word"];
        tds[i].style.backgroundColor = calendar.colors["cur_bg"];
        tds[i].onmouseover = function () {
          this.style.backgroundColor = calendar.colors["td_bg_over"];
        }
        tds[i].onmouseout = function () {
          this.style.backgroundColor = calendar.colors["cur_bg"];
        }
        //continue; //若不想当天单元格的背景被下面的覆盖，请取消注释 →  2006-12-03 寒羽枫添加
      }//end if
      //设置已被选择的日期单元格背影色 2006-12-03 寒羽枫添加
   // if (calendar.dateControl != null && calendar.dateControl.value == new Date(calendar.date.getFullYear(),
      //             calendar.date.getMonth(),
      //             dateArray[i]).format(calendar.dateFormatStyle)) {

      //mazq 修改 在带有时间的条件下，设置已被选择的日期单元格 2007-4-5 
      if (calendar.dateControl != null && calendar.dateControl.value == new Date(calendar.date.getFullYear(),
                   calendar.date.getMonth(),
                   dateArray[i], c_hours.value, c_minutes.value, c_seconds.value).format(calendar.dateFormatStyle)) {
        tds[i].style.backgroundColor = calendar.colors["sel_bg"];
        tds[i].onmouseover = function () {
          this.style.backgroundColor = calendar.colors["td_bg_over"];
        }
        tds[i].onmouseout = function () {
          this.style.backgroundColor = calendar.colors["sel_bg"];
        }
      }
    }
  }
}
//根据年、月得到月视图数据(数组形式)
Calendar.prototype.getMonthViewArray = function (y, m) {
  var mvArray = [];
  var dayOfFirstDay = new Date(y, m, 1).getDay();
  var daysOfMonth = new Date(y, m + 1, 0).getDate();
  for (var i = 0; i < 42; i++) {
    mvArray[i] = "&nbsp;";
  }
  for (var i = 0; i < daysOfMonth; i++){
    mvArray[i + dayOfFirstDay] = i + 1;
  }
  return mvArray;
}

//扩展 document.getElementById(id) 多浏览器兼容性 from meizz tree source
Calendar.prototype.getElementById = function(id){
  if (typeof(id) != "string" || id == "") return null;
  if (document.getElementById) return document.getElementById(id);
  if (document.all) return document.all(id);
  try {return eval(id);} catch(e){ return null;}
}

//扩展 object.getElementsByTagName(tagName)
Calendar.prototype.getElementsByTagName = function(object, tagName){
  if (document.getElementsByTagName) return document.getElementsByTagName(tagName);
  if (document.all) return document.all.tags(tagName);
}

//取得HTML控件绝对位置
Calendar.prototype.getAbsPoint = function (e){
  var x = e.offsetLeft;
  var y = e.offsetTop;
  while(e = e.offsetParent){
    x += e.offsetLeft;
    y += e.offsetTop;
  }
  return {"x": x, "y": y};
}

//显示日历
Calendar.prototype.show = function (dateObj, popControl) {
  if (dateObj == null){
    throw new Error("arguments[0] is necessary")
  }
  this.dateControl = dateObj;
  
  //if (dateObj.value.length > 0){
  //this.date = new Date(dateObj.value.toDate());
  //this.date = new Date(dateObj.value.toDate(this.dateFormatStyle));//由寒羽枫修改，带入用户指定的 style  
  this.date = (dateObj.value.length > 0) ? new Date(dateObj.value.toDate(this.dateFormatStyle)) : new Date() ;//2006-12-03 寒羽枫添加 → 若为空则显示当前月份
  this.year = this.date.getFullYear();
  this.month = this.date.getMonth();
  this.changeSelect();
  this.bindData();

  //以下 mazq 2007-4-4 添加
  this.hours = this.date.getHours();
  this.minutes = this.date.getMinutes();
  this.seconds = this.date.getSeconds();

  c_hours.value = this.hours;
  c_minutes.value = this.minutes;
  c_seconds.value = this.seconds;
 //以上 mazq 2007-4-4 添加
  //}
  if (popControl == null){
    popControl = dateObj;
  }
  var xy = this.getAbsPoint(popControl);
  this.panel.style.left = xy.x -25 + "px";
  this.panel.style.top = (xy.y + dateObj.offsetHeight) + "px";
  
  //由寒羽枫 2006-06-25 修改 → 把 visibility 变为 display，并添加失去焦点的事件
  //this.setDisplayStyle("select", "hidden");
  //this.panel.style.visibility = "visible";
  //this.container.style.visibility = "visible";
  this.panel.style.display = "";
  this.container.style.display = "";
  
  dateObj.onblur = function(){calendar.onblur();}
  this.container.onmouseover = function(){isFocus=true;}
  this.container.onmouseout = function(){isFocus=false;}
}

//隐藏日历
Calendar.prototype.hide = function() {
  //this.setDisplayStyle("select", "visible");
  //this.panel.style.visibility = "hidden";
  //this.container.style.visibility = "hidden";
  this.panel.style.display = "none";
  this.container.style.display = "none";
  isFocus=false;
}

//焦点转移时隐藏日历 → 由寒羽枫 2006-06-25 添加
Calendar.prototype.onblur = function() {
    if(!isFocus){this.hide();}
}

//以下由寒羽枫 2006-06-25 修改 → 用<iframe> 遮住 IE 的下拉框
/**//**//**//*
//设置控件显示或隐藏
Calendar.prototype.setDisplayStyle = function(tagName, style) {
  var tags = this.getElementsByTagName(null, tagName)
  for(var i = 0; i < tags.length; i++) {
    if (tagName.toLowerCase() == "select" &&
       (tags[i].name == "calendarYear" ||
      tags[i].name == "calendarMonth")){
      continue;
    }
    //tags[i].style.visibility = style;
    tags[i].style.display = style;
  }
}
*/
//document.write('<div id="ContainerPanel" style="visibility:hidden"><div id="calendarPanel" style="position: absolute;visibility: hidden;z-index: 9999;');
document.write('<div id="ContainerPanel" style="display:none;z-index:20000"><div id="calendarPanel" style="position: absolute;display: none;z-index: 22222222222222222222222222;');
document.write('background-color: #FFFFFF;border: 1px solid #CCCCCC;width:175px;font-size:10px;"></div>');
if(document.all)
{
document.write('<iframe style="position:absolute;z-index:2000;width:expression(this.previousSibling.offsetWidth);');
document.write('height:expression(this.previousSibling.offsetHeight);');
document.write('left:expression(this.previousSibling.offsetLeft);top:expression(this.previousSibling.offsetTop);');
document.write('display:expression(this.previousSibling.style.display);" scrolling="no" frameborder="no"></iframe>');
}
document.write('</div>');
//var calendar = new Calendar();  //此句被 寒羽枫注释，否则 IE 将报错
//调用calendar.show(dateControl, popControl);

//以下 mazq 2007-4-4 添加
Calendar.prototype.play = function()
 {
  //this.timer = setInterval("calendarTime.playback()",1000);
 }
 Calendar.prototype.formatTime = function(sTime)
 {
  sTime = ("0"+sTime);
  return sTime.substr(sTime.length-2);
 }
Calendar.prototype.playback = function()
 {
  var objDate = new Date();
  var arrDate = [objDate.getHours(),objDate.getMinutes(),objDate.getSeconds()];
  var objMinute = this.getElementById("txtcalendarTime");
  for (var i=0;i<objMinute.length;i++)
  {
   objMinute[i].value = this.formatTime(arrDate[i])
  }
 }
 Calendar.prototype.prevent = function(obj)
 {
  clearInterval(this.timer);
  this.setFocusObj(obj);
  var value = parseInt(obj.value,10);
  var radix = parseInt(obj.radix,10)-1;
  if (obj.value>radix||obj.value<0)
  {
   obj.value = obj.value.substr(0,1);
  }
 }
 Calendar.prototype.controlTime = function(cmd)
 {
  event.cancelBubble = true;
  if (!this.fObj) return;
  clearInterval(this.timer);
  var cmd = event.srcElement.innerText=="5"?true:false;
  var i = parseInt(this.fObj.value,10);
  var radix = parseInt(this.fObj.radix,10)-1;
  if (i==radix&&cmd)
  {
   i = 0;
  }
  else if (i==0&&!cmd)
  {
   i = radix;
  }
  else
  {
   cmd?i++:i--;
  }
  this.fObj.value = this.formatTime(i);
  this.fObj.select();
 }
 Calendar.prototype.setTime = function(obj)
 {
  obj.value = this.formatTime(obj.value);
 }
 Calendar.prototype.setFocusObj = function(obj)
 {
  eval("c_up").disabled = eval("c_down").disabled = false;
  this.fObj = obj;
 }
 Calendar.prototype.getTime = function()
 {
  var arrTime = new Array(2);
  for (var i=0;i<this.getElementById("txtcalendarTime").length;i++)
  {
   arrTime[i] = this.getElementById("txtcalendarTime")[i].value;
  }
  return arrTime.join(":");
 }
//以上 mazq 2007-4-4 添加
//-->
