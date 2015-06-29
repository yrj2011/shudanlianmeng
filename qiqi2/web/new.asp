<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统**************************************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=stupname%>-<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.f12bt{
	font-size: 12px;
	color: #024E68;
	line-height: 20px;
	text-decoration: none;
	font-weight: bold;
}
-->
</style>

</head>

<body>  
<!--#include file=jieducm_top.asp-->
<div align="center">
<TABLE cellSpacing=0 cellPadding=0 width=960 align=center border=0>
  

  <tr>
    <td valign="top" align="left"><table width="970" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:10px 0 10px 0;">
      <tr>
        <td colspan="3" valign="top">		</td>
        </tr>
      <tr>
        <td width="285" valign="top"><table width="260" border="0"  cellpadding="0" cellspacing="0">
          <tr>
            <td><TABLE cellSpacing=0 cellPadding=0 width=260 border=0>
        <TBODY>
        <TR>
          <TD width=20><IMG height=32 
            src="images/Top_10.gif" 
            width=20></TD>
          <TD width=10><IMG height=32 
            src="images/Top_9.gif" 
            width=10></TD>
          <TD class=K_mttitle width=95 
          background=images/Top_12.gif>网站公告</TD>
          <TD width=10><IMG height=32 
            src="images/Top_13.gif" 
            width=10></TD>
          <TD width=119 
          background=images/Top_11.gif></TD>
          <TD width=6><IMG height=32 
            src="images/Top_14.gif" 
            width=6></TD></TR>
        <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(2,10,15)%></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
          </tr>
		  <tr height="10"><td></td></tr>
          <tr>
            <td><TABLE cellSpacing=0 cellPadding=0 width=260 border=0>
        <TBODY>
        <TR>
          <TD width=20><IMG height=32 
            src="images/Top_10.gif" 
            width=20></TD>
          <TD width=10><IMG height=32 
            src="images/Top_9.gif" 
            width=10></TD>
          <TD class=K_mttitle width=95 
          background=images/Top_12.gif>推荐文章</TD>
          <TD width=10><IMG height=32 
            src="images/Top_13.gif" 
            width=10></TD>
          <TD width=119 
          background=images/Top_11.gif></TD>
          <TD width=6><IMG height=32 
            src="images/Top_14.gif" 
            width=6></TD></TR>
        <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call elitenew(10,15)%></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
          </tr>
        </table></td>
        
        <td width="715" valign="top"><table width="699" border="0" align="right" cellpadding="0" cellspacing="0">
          <tr>
            <td width="349" valign="top"><table width="334" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24" background="images/lun-1.gif"><table width="310" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">新手入门</td>
                        <td width="249" align="right"><a href="newmore.asp?action=90">更多>></a></td>
                      </tr>
                  </table></td>
                </tr>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(90,5,15)%></TBODY></TABLE></TD></TR>
            </table></td>
            <td width="350" valign="top"><table width="334" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24" background="images/lun-1.gif"><table width="310" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">刷客必读</td>
                        <td width="249" align="right"><a href="newmore.asp?action=91">更多>></a></td>
                      </tr>
                  </table></td>
                </tr>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(91,5,15)%></TBODY></TABLE></TD></TR>
            </table></td>
          </tr>
          <tr>
            <td valign="top">&nbsp;</td>
            <td valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top"><table width="334" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24" background="images/lun-1.gif"><table width="310" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">卖家必读</td>
                        <td width="249" align="right"><a href="newmore.asp?action=92">更多>></a></td>
                      </tr>
                  </table></td>
                </tr>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(92,5,15)%></TBODY></TABLE></TD></TR>
            </table></td>
            <td valign="top"><table width="334" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24" background="images/lun-1.gif"><table width="310" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">买家必读</td>
                        <td width="249" align="right"><a href="newmore.asp?action=93">更多>></a></td>
                      </tr>
                  </table></td>
                </tr>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(93,5,15)%></TBODY></TABLE></TD></TR>
            </table></td>
          </tr>
          <tr>
            <td valign="top">&nbsp;</td>
            <td valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top"><table width="334" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24" background="images/lun-1.gif"><table width="310" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">店铺推广</td>
                        <td width="249" align="right"><a href="newmore.asp?action=94">更多>></a></td>
                      </tr>
                  </table></td>
                </tr>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(94,5,15)%></TBODY></TABLE></TD></TR>
            </table></td>
            <td valign="top"><table width="334" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24" background="images/lun-1.gif"><table width="310" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">赚钱窍门</td>
                        <td width="249" align="right"><a href="newmore.asp?action=95">更多>></a></td>
                      </tr>
                  </table></td>
                </tr>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(95,5,15)%></TBODY></TABLE></TD></TR>
            </table></td>
          </tr>
          <tr>
            <td valign="top">&nbsp;</td>
            <td valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top"><table width="334" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24" background="images/lun-1.gif"><table width="310" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">网络营销 </td>
                        <td width="249" align="right"><a href="newmore.asp?action=96">更多>></a></td>
                      </tr>
                  </table></td>
                </tr>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(96,5,15)%></TBODY></TABLE></TD></TR>
            </table></td>
            <td valign="top"><table width="334" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24" background="images/lun-1.gif"><table width="310" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">视频教室</td>
                        <td width="249" align="right"><a href="newmore.asp?action=97">更多>></a></td>
                      </tr>
                  </table></td>
                </tr>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(97,5,15)%></TBODY></TABLE></TD></TR>
            </table></td>
          </tr>
        </table></td>
      </tr>
      
    </table></td>
  </tr>
  
 
</table>
</div>
<!--#include file=jieducm_bottom.asp-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html>
