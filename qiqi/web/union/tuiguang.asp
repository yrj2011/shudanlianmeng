<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!--#include file="../config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>会员中心-<%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="union.css" type="text/css" rel="stylesheet" />
<LINK href="Css.css" type=text/css rel=stylesheet>
<style type="text/css">

.STYLE5 {font-size: 25px}
.STYLE6 {
	font-size: 14px;
	font-weight: bold;
}

-->
</style>
</head>

<body>
<!--#include file=../inc/jieducm_top.asp-->
<table width="960" border="0" align="center">
  <tr>
    <td width="209">&nbsp;</td>
  </tr>
</table>
<table width="960" border="0" align="center">
  <tr>
    <td width="209">快速获得推广链接，直接进行推广：</td>
    <td width="741">
	  <script language=javascript>
    // 自动 COPY 代码开始
    function MM_findObj(n, d) { //v4.0
    var p,i,x; if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
    if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
    for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
    if(!x && document.getElementById) x=document.getElementById(n); return x;
    }
    function JM_cc(ob){
    var obj=MM_findObj(ob); if (obj) {
    obj.select();js=obj.createTextRange();js.execCommand("Copy");}
    alert("复制成功！");
    }
    
    </script>
	<label>
     <input name="page_url" type="text" value="<%=stupurl%>/register.asp?promotion=<%=uid%>" size="80" style="height:20px;"/> 
      <input type="submit" name="Submit"  style="height:30px; width:50px;"  value="复制" onClick=JM_cc("page_url") />	  
    </label></td>
  </tr>
</table>
<table width="960" height="140" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="../img/tui.jpg" width="960" height="176" /></td>
  </tr>
</table>
<table width="960" height="40" border="0" align="center" cellpadding="0" cellspacing="0"  >
  <tbody>
    <tr>
      <th width="157"><table width="303" height="31" border="0" cellpadding="0" cellspacing="1" bgcolor="#D0DBE3" >
        <tr>
          <td bgcolor="#F0F7FB"><div align="center">文字推广</div></td>
          <td bgcolor="#F0F7FB"><div align="center">图片推广</div></td>
          <td bgcolor="#F0F7FB"><div align="center">邮件推广</div></td>
        </tr>
      </table></th>
    </tr>
  </tbody>
</table>
<table width="960" border="0" align="center" cellpadding="10" cellspacing="0" bgcolor="#F0F7FB"  class="border_hui margin-bottom">
  <tbody>
    <tr>
      <th width="157"><img height="22" alt="" src="../img/jieducm_union_01.gif" 
      width="25" align="absmiddle" /><a>文字代码&nbsp;<img 
      src="../img/3jiaoxia.gif" border="0" /></a> </th>
      <td width="595">适合发给QQ/MSN好友、个人坛签名或个人博客</td>
    </tr>
    <tr>
      <td colspan="2"><table cellpadding="10" class="getCode" id="type_a">
        <tbody>
          <tr>
            <td class="em" width="31"><em>1</em></td>
            <td width="453">样式一</td>
            <td width="252">发给你的QQ/MSN好友</td>
          </tr>
          <tr>
            <td class="em"></td>
            <td><input name="Input" id="url001" value="<%=stupurl%>/register.asp?promotion=<%=uid%>" size="46" />
            </td>
            <td><input name="submit" type="submit" class="button" onClick=JM_cc("url001") value="复制" />
			
			</td>
          </tr>
          <tr>
            <td colspan="3"><hr />
            </td>
          </tr>
          <tr>
            <td class="em"><em>2</em></td>
            <td>样式二</td>
            <td>适合空间/博客友情链接 或论坛签名</td>
          </tr>
          <tr>
            <td class="em"></td>
            <td>HTML:
                  <input name="Input" id="url002" 
            value="&lt;a href='<%=stupurl%>/register.asp?promotion=<%=uid%>' target='_blank'&gt;<%=uid%>&lt;/a&gt;" size="46" />
                  <br />
                <br />
              UBB:&nbsp;&nbsp;
              <input name="Input" 
            id="url0022" 
            value="[url=<%=stupurl%>/register.asp?promotion=<%=uid%>]<%=uid%>[/url]" size="46" />
                  <br />
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
            <td valign="top"><input name="submit" type="submit" class="button" onClick=JM_cc("url002") value="复制" />
                  <br />
              <input name="submit" type="submit" class="button" onClick=JM_cc("url0022") value="复制" /></td>
          </tr>
          <tr>
            <td colspan="3"><hr />
            </td>
          </tr>
          <tr>
            <td class="em"><em>3</em></td>
            <td>样式三</td>
            <td>适合空间/博客友情链接 或论坛签名</td>
          </tr>
          <tr>
            <td class="em"></td>
            <td>HTML:
                  <input name="Input" id="url003" 
            value="&lt;a href='<%=stupurl%>/register.asp?promotion=<%=uid%>' target='_blank'&gt;<%=uid%>&lt;/a&gt;" size="46" />
                  <br />
                  <br />
              UBB:&nbsp;&nbsp;
              <input name="Input" id="url0032" 
            value="[url=<%=stupurl%>/register.asp?promotion=<%=uid%>]<%=uid%>[/url]" size="46" />
                  <br />
              &nbsp;&nbsp;&nbsp;</td>
            <td valign="top"><input name="submit" type="submit" class="button" onClick=JM_cc("url003") value="复制" />
                  <br />
              <input name="submit" type="submit" class="button" onClick=JM_cc("url0032") value="复制" /></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>
