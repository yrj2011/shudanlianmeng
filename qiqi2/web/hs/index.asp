<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="jieducm_ad.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<%
'�뱣����������Ϣ,�������������Ӱ������ٶ�!
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
%>
<HTML lang=en><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta  name="description" content="<%=stupdesc%>"/>
<meta  name="keywords" content="<%=stupkey%>"/>
<LINK href="hs/main.css" type=text/css rel=stylesheet>
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<link rel="stylesheet" id='skin' type="text/css" href="skin/qq/ymPrompt.css" />
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
<script type="text/javascript" src="js/ymPrompt.js"></script>
</HEAD>
<BODY  onload=document.form.name.focus()> 
<script language=javascript>
function CheckForm()
{
	if(document.form.name.value=="")
	{
		alert("�������û�����");
		document.form.name.focus();
		return false;
	}
	if(document.form.pass.value == "")
	{
		alert("���������룡");
		document.form.pass.focus();
		return false;
	}
	if (document.form.CheckCode)
	{
		if (document.form.CheckCode.value==""){
		   alert ("������������֤�룡");
		   document.form.CheckCode.focus();
		   return(false);
		}
	}
}
<%if stuppopup=1 then%>
	ymPrompt.alert({message:'<%=stupgonggao%>',title:'<%=stupkouhao%>!',height:250,width:400,fixPosition:true,dragOut:false});
		window.onload=function(){
			var o=document.getElementById('chgSkin');
			var css=document.getElementById('skin');
			o.selectedIndex=0;
			o.onchange=function(){
				css.href='skin/'+this.options[this.selectedIndex].value+'/ymPrompt.css';
			}
		}
		function json2str(o){
			var arr=[];
			var fmt=function(s){
				return /^(string|number)$/.test(typeof s)?"'"+s+"'":s;
			}
			for(var i in o) arr.push(i+':'+fmt(o[i]));
			return '{<br>&nbsp;&nbsp;'+arr.join(',<br>&nbsp;&nbsp;')+'&nbsp;&nbsp;<br>}';
		}
<%end if%>
</script>
<!--#include file=jieducm_top.asp-->
<%
sql="SELECT  sum(price) as jieducm_total FROM "&jieducm&"_zhongxin"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_total=rs("jieducm_total")
end if 

sql="SELECT  count(*) as jieducm_count FROM "&jieducm&"_zhongxin"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_count=rs("jieducm_count")
end if 
%>
<DIV class=page>
<DIV class=broadcast>
<DIV class=index-number style="OVERFLOW: hidden; WIDTH: 460px">
<span style="font-size:19px; color:#ff5e10" ><img src="images/jieducm_list.jpg" width="28" height="28" border="0">ƽָ̨��</span><SPAN>���ֵ����<EM><%=jieducm_total%></EM>Ԫ</SPAN><SPAN>����������<EM><%=jieducm_count%></EM>��</SPAN>
</DIV>
<DIV class=real-time-info style="TEXT-ALIGN: center">
<MARQUEE scrollAmount=1 direction=up width=450 height=20>
<%
Sql = "select top 20 * from "&jieducm&"_zhongxin where  classid<>'6' order by id desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
Response.Write("������Ϣ")
Else
Do While Not Rs.Eof		
%>
<P style="MARGIN-BOTTOM: 10px"><EM><%=left(rs("username"),1)%>***<%=right(rs("username"),1)%> </EM>�� <SPAN><%=rs("now")%> 
</SPAN><EM>������ </EM>��һ��������Ϊ<EM><%=rs("price")%></EM>Ԫ������</P>
<%Rs.MoveNext
Loop
End IF%>
</MARQUEE></DIV></DIV></DIV>
<div class="main" style="height:226px;">
<div class="login" style="overflow:hidden;">
<div class="logint"><img src="hs/loginletter.gif" alt="ƽ̨���¹���" /></div>
<div class="loginm1">
         <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(2,10,15)%></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
          </tr><div class="li2"></div>

</div>
</div>

<div class="mainads"><div class="centeradsm">
  <!--#include file="jieducm_flash.asp"-->
</div> 
</div>
</div>
<div class="tj">
<div class="middleBar">
  <div class="left"></div>
  <div class="right"></div>
  <div class="middle">
  
      <div class="howTo">
          <div class="title"><h2>�Ա�ˢ��������</h2><em>С��ʾ����ʵ�ɿ���<b>�Ա�ˢ����</b>��<b>�Ա�ˢ��</b>��<b>�Ա�ˢ����</b>��Ȼ�죬��ǧ��Ҫ̰��Ŷ��</em></div>
          <div class="cont">
              <ol>
                <li class="step1">ע�ἤ��</li>
                <li class="step2">�������</li>
                <li class="step3">�����<span>����ֵ</span></li>
                <li class="step4">��������׬ȡ������<span>ʹ�ô��򷢲���</span></li>
                <li class="step5">������������������</li>
                <li class="step6">˫����������</li>
              </ol>
          </div>
      </div>  
  </div>
</div></div>
<div class="main1">
<div class="inleft">
  <div class="helpnews">
<div class="helpnewst"><h3><table width="220" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="60" class="f12bt">��������</td>
                        <td width="170" align="right"><a href="newmore.asp?action=90">����>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpnewsm"><ul>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(90,8,15)%></TBODY></TABLE></TD></TR>
</ul>

</div>
</div></div>
<div class="inright">
<div class="helpteach" style="margin-top:0;">
<div class="helpteacht"><h3><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">���ұض�</td>
                        <td width="249" align="right"><a href="newmore.asp?action=92">����>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpteachm">
<ul style="background:url(hs/jc1b.gif) no-repeat 220px center">
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(92,8,15)%></TBODY></TABLE></TD></TR>
</ul>
</div>
</div>
<div class="helpteach" style="margin-left:7px;margin-top:0;">
<div class="helpteacht"><h3><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">��ұض�</td>
                        <td width="249" align="right"><a href="newmore.asp?action=93">����>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpteachm">
<ul style="background:url(hs/jc2b.gif) no-repeat 220px center">
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(93,8,15)%></TBODY></TABLE></TD></TR>
</ul>
</div>
</div></div>
</div>
<div class="main1">
<div class="inleft">
  <div class="helpnews">
<div class="helpnewst"><h3><table width="220" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="60" class="f12bt">ˢ�ͱض�</td>
                        <td width="170" align="right"><a href="newmore.asp?action=91">����>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpnewsm"><ul>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(91,8,15)%></TBODY></TABLE></TD></TR>
</ul>

</div>
</div></div>
<div class="inright">
<div class="helpteach" style="margin-top:0;">
<div class="helpteacht"><h3><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">׬Ǯ����</td>
                        <td width="249" align="right"><a href="newmore.asp?action=95">����>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpteachm">
<ul style="background:url(hs/jc1b.gif) no-repeat 220px center">
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(95,8,15)%></TBODY></TABLE></TD></TR>
</ul>
</div>
</div>
<div class="helpteach" style="margin-left:7px;margin-top:0;">
<div class="helpteacht"><h3><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">�����ƹ�</td>
                        <td width="249" align="right"><a href="newmore.asp?action=93">����>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpteachm">
<ul style="background:url(hs/jc2b.gif) no-repeat 220px center">
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(94,8,15)%></TBODY></TABLE></TD></TR>
</ul>
</div>
</div></div>
</div>
<!--#include file=jieducm_bottom.asp-->
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>