<%@language=vbscript codepage=936 %>

<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
 <%
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
action=request("action")
if action="ok" then

Set rs1=server.createobject("ADODB.RECORDSET")
rs1.open "select * from "&jieducm&"_stup",conn,3,3
if not (rs1.eof) then
rs1("name")=request.Form("SiteName")
rs1("title")=request.Form("SiteTitle")
rs1("key")=request.Form("SiteKey")
rs1("desc")=request.Form("Sitedes")
rs1("gonggao")=request.Form("gonggao")
rs1("kouhao")=request.Form("kouhao")
rs1("url")=request.Form("SiteUrl")
rs1("logo")=request.Form("LogoUrl")
rs1("address")=request.Form("BannerUrl")
rs1("qq")=request.Form("qq")
rs1("qq2")=request.Form("qq2")
rs1("qq3")=request.Form("qq3")
rs1("qq4")=request.Form("qq4")
rs1("qq5")=request.Form("qq5")
rs1("qq6")=request.Form("qq6")
rs1("zfb")=request.Form("WebmasterName")
rs1("cft")=request.Form("cft")
rs1("email")=request.Form("WebmasterEmail")
rs1("phone")=request.Form("phone")
rs1("tel")=request.Form("tel")
rs1("icp")=request.Form("icp")
rs1("copy")=request.Form("Copyright")
rs1("fa")=request.Form("fa")
rs1("jifei")=request.Form("jifei")
rs1("kuan")=request.Form("kuan")
rs1("kuanvip")=request.Form("kuanvip")
rs1("kou")=request.Form("kou")
rs1("geshu")=request.Form("geshu")
rs1("zhang")=request.Form("zhang")
rs1("ping")=request.Form("ping")
rs1("zhu")=request.Form("zhu")
rs1("tjrjifei")=request.Form("tjrjifei")
rs1("tjrzhu")=request.Form("tjrzhu")
rs1("vipjifei")=request.Form("vipjifei")
rs1("fajifei")=request.Form("fajifei")
rs1("shouc")=request.Form("shouc")
rs1("shouf")=request.Form("shouf")
rs1("shouf2")=request.Form("shouf2")
rs1("shouxu")=request.Form("shouxu")
rs1("wu")=request.Form("wu")
rs1("shi")=request.Form("shi")
rs1("wushi")=request.Form("wushi")
rs1("yibai")=request.Form("yibai")
rs1("wubai")=request.Form("wubai")
rs1("ka")=request.Form("ka")
rs1("jifeidi")=request.Form("jifeidi")
rs1("jifeigao")=request.Form("jifeigao")
rs1("alllogin")=request.Form("alllogin")
rs1("fhc")=request.Form("fhc")
rs1("fhcshu")=request.Form("fhcshu")
rs1("quxin")=request.Form("quxin")
rs1("qushou")=request.Form("qushou")
rs1("quvip")=request.Form("quvip")
rs1("buyhao")=request.Form("buyhao")
rs1("xinshoufa")=request.Form("xinshoufa")
rs1("shouchangfa")=request.Form("shouchangfa")
rs1("vipfa")=request.Form("vipfa")
rs1("enchfa")=request.Form("enchfa")
rs1("jifeibig")=request.Form("jifeibig")
rs1("songfa")=request.Form("songfa")
rs1("songshou")=request.Form("songshou")
rs1("songji")=request.Form("songji")
rs1("car")=request.Form("car")
rs1("wangyin")=request.Form("wangyin")
rs1("alipay")=request.Form("alipay")
rs1("login")=request.Form("login2")
rs1("CheckCode")=request.Form("CheckCode")
rs1("isgive")=request.Form("isgive")
rs1("popup")=request.Form("popup")
rs1("register_taobao")=request.Form("register_taobao")
rs1("register_pai")=request.Form("register_pai")
rs1("register_you")=request.Form("register_you")
rs1("MailServer")=request.Form("MailServer")
rs1("MailServerUserName")=request.Form("MailServerUserName")
rs1("MailServerPassWord")=request.Form("MailServerPassWord")
rs1("MailDomain")=request.Form("MailDomain")
rs1("isemail")=request.Form("isemail")
rs1("ismessage")=request.Form("ismessage")
rs1("emailcontent")=request.Form("emailcontent")
rs1("messagecontent")=request.Form("messagecontent")
rs1("emailactive")=request.Form("emailactive")
rs1("puuser")=request.Form("puuser")
rs1("vipuser")=request.Form("vipuser")
rs1("tribestart")=request.Form("tribestart")
rs1("phonestart")=request.Form("phonestart")
rs1("jieweiok")=request.Form("jieweiok")
rs1("faweiok")=request.Form("faweiok")
rs1("invite")=request.Form("invite")
rs1("exitauto")=request.Form("exitauto")
rs1("vip_car_jieducm")=request.Form("vip_car_jieducm")
rs1("cang")=request.Form("cang")
rs1("kefu_pic")=request.Form("kefu_pic")
rs1("jieducm_gonggao")=request.Form("jieducm_gonggao")
rs1("jieducm_MerId")=request.Form("jieducm_MerId")
rs1("jieducm_MerKey")=request.Form("jieducm_MerKey")
rs1("yibao")=request.Form("yibao")
rs1("payis")=request.Form("payis")
rs1("payjifei")=request.Form("payjifei")
rs1("paynum")=request.Form("paynum")
rs1("vippaynum")=request.Form("vippaynum")
rs1("pupaynum")=request.Form("pupaynum")
rs1("RndQueryNum")=request.Form("RndQueryNum")
rs1("DefQueryNum")=request.Form("DefQueryNum")
rs1.update
rs1.close
set rs1=nothing
Response.Write("<script>alert('�޸ĳɹ�!');window.location.href='admin_siteconfig.asp';</script>")
end if
end if
	  
			  	Sql = "select* from "&jieducm&"_stup"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("������Ϣ")
				Else
  %>
<html>
<head>
<title>��վ��Ϣ����ѡ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" Class="border">
    <tr> 
     <td height="22" colspan=2 align=center class="title"><a name="Top"></a><b>�� վ �� ��</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="70" height="30"><strong>��������</strong></td>
      
    <td height="30"><a href="#SiteInfo">��վ��Ϣ����</a>| <a href="#Email">�ʼ�������ѡ��</a></td>
    </tr>
  </table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form method="POST" action="Admin_SiteConfig.asp?action=ok" id="frm1" name="frm1"><tr><td height="25" class="tdbg"></td>
  <td width="942" height="25" class="tdbg">     </td>
    <tr>
      <td height="22" colspan="2" class="title"><strong>��վ��Ϣ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong><a href="#Top"><font color="#0066CC">����&nbsp;����</font></a></td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>��վ���ƣ�</strong></td>
      <td width="942" height="25" class="tdbg"><input name="SiteName" type="text" id="SiteName" value="<%=rs("name")%>" size="40" maxlength="150">      </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>��վ���⣺</strong></td>
      <td width="942" height="25" class="tdbg"><input name="SiteTitle" type="text" id="SiteTitle" value="<%=rs("title")%>" size="40" maxlength="250">      </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>��վ�ؼ��֣�</strong><br>
        ������������������������վ�Ĺؼ�����<br>
        ÿ���ؼ����á�|���ŷָ�</td>
      <td width="942" height="25" class="tdbg"><input name="SiteKey" type="text" id="SiteKey" value="<%=rs("key")%>" size="40" maxlength="250">      </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>��վ������</strong><br>
        ����������������˵������վ����Ҫ����<br>
        <font color="#FF6600">�������벻Ҫ��Ӣ�ĵĶ���</font></td>
      <td width="942" height="25" class="tdbg"><input name="Sitedes" type="text" id="Sitedes" value="<%=rs("desc")%>" size="40" maxlength="250">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>ϵͳ����״̬��</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="radio" name="invite" value="1"  <% if rs("invite")=1 then%> checked <% end if%>>
        ��������      
        &nbsp;&nbsp;&nbsp;  <input type="radio" name="invite" value="2"  <% if rs("invite")=2 then%> checked <% end if%>>
        ��ͣά��
    </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�Ƿ�ر�ˢ����</strong></td>
      <td height="25" class="tdbg">
	    <input name="quxin" type="checkbox" id="quxin" value="1" <%if rs("quxin")=1 then%> checked <%end if%>>
        �Ա��� 
          <input name="qushou" type="checkbox" id="qushou" value="1" <%if rs("qushou")=1 then%> checked <%end if%>>
          ������          </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�Ƿ�ر�ע�᣺</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="zhu" id="zhu" value="1"  <% if rs("zhu")="1" then%> checked <% end if%> >
        �ر�ע�� 
          <input name="zhu" type="radio" id="zhu" value="2" <% if rs("zhu")="2" then%> checked <% end if%> >
          ����ע��</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>ע���Ա�Ƿ���Ҫ�ʼ����</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="radio" name="emailactive" value="1"  <% if rs("emailactive")="1" then%> checked <% end if%> >��Ҫ
        <input type="radio" name="emailactive" value="2" <% if rs("emailactive")="2" then%> checked <% end if%>>����Ҫ
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>ע���Ա�Ƿ���Ҫ������֤�룺</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="CheckCode" id="CheckCode" value="1"  <% if rs("CheckCode")="1" then%> checked <% end if%> >
�ر���֤��
  <input name="CheckCode" type="radio" id="CheckCode" value="2" <% if rs("CheckCode")="2" then%> checked <% end if%> >
������֤��</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>ע���Ա�Ƿ���Ҫ��ˣ�</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="login2" id="login2" value="1"  <% if rs("login")="1" then%> checked <% end if%> >
        �ر���� 
          <input name="login2" type="radio" id="login2" value="0" <% if rs("login")="0" then%> checked <% end if%> >
          �������</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��Ա�Ƿ������벿����ܽӷ�����</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="tribestart" id="tribestart" value="1"  <% if rs("tribestart")="1" then%> checked <% end if%> >
��Ҫ
  <input name="tribestart" type="radio" id="tribestart" value="0" <% if rs("tribestart")="0" then%> checked <% end if%> >
����Ҫ</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��Ա�Ƿ������ֻ����ܽӷ�����</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="phonestart" id="phonestart" value="1"  <% if rs("phonestart")="1" then%> checked <% end if%> >
��Ҫ
  <input name="phonestart" type="radio" id="phonestart" value="0" <% if rs("phonestart")="0" then%> checked <% end if%> >
����Ҫ</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���û�ע�Ὺͨ��</strong></td>
      <td height="25" class="tdbg">
	  <input name="register_taobao" type="checkbox" id="register_taobao" value="1" <%if rs("register_taobao")=1 then%> checked <%end if%>>
�Ա���
  <input name="register_pai" type="checkbox" id="register_pai" value="1" <%if rs("register_pai")=1 then%> checked <%end if%>>
������</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>ע���Ա�Ƿ�������ͷ����㣺</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="isgive" id="isgive" value="1"  <% if rs("isgive")="1" then%> checked <% end if%> >
������
  <input name="isgive" type="radio" id="isgive" value="0" <% if rs("isgive")="0" then%> checked <% end if%> > 
  ����</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�Ƿ�رճ�ֵ���ֶ����</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="ka" id="ka" value="1"  <% if rs("ka")="1" then%> checked <% end if%> >
        �ر��ֶ����� 
          <input name="ka" type="radio" id="ka" value="2" <% if rs("ka")="2" then%> checked <% end if%> >
          �����ֶ�����</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�Ƿ�رճ�ֵ����ֵ��</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="car" id="car" value="0" <% if rs("car")="0" then%> checked <% end if%>>
        �رճ�ֵ����ֵ
        <input name="car" type="radio" id="car" value="1" <% if rs("car")="1" then%> checked <% end if%>>
        ������ֵ����ֵ</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�Ƿ�ر��������߳�ֵ��</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="wangyin" id="wangyin" value="0" <% if rs("wangyin")="0" then%> checked <% end if%>>
        �ر��������߳�ֵ
        <input type="radio" name="wangyin" id="wangyin" value="1" <% if rs("wangyin")="1" then%> checked <% end if%>>
        �����������߳�ֵ</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�Ƿ�ر�֧�����ӿڣ�</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="alipay" id="alipay" value="0" <% if rs("alipay")="0" then%> checked <% end if%>>
        �ر�֧�����ӿ�
          <input type="radio" name="alipay" id="alipay" value="1" <% if rs("alipay")="1" then%> checked <% end if%>>
        ����֧�����ӿ�</td>
    </tr>
	
	<tr>
      <td height="25" class="tdbg"><strong>�Ƿ�ر��ױ�֧����</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="yibao" id="yibao" value="0" <% if rs("yibao")="0" then%> checked <% end if%>>
        �ر��ױ�֧��
          <input type="radio" name="yibao" id="yibao" value="1" <% if rs("yibao")="1" then%> checked <% end if%>>
        �����ױ�֧��</td>
    </tr>
	
	
    <tr>
      <td height="25" class="tdbg"><strong>�Ƿ�ر������û���¼��</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="alllogin" id="alllogin" value="1" <% if rs("alllogin")="1" then%> checked <% end if%>>
        ���������û���¼ 
          <input type="radio" name="alllogin" id="alllogin" value="0" <% if rs("alllogin")="0" then%> checked <% end if%>>
          �ָ������û���¼</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>����������һ��ֽ�</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="fhc" id="fhc" value="1" <% if rs("fhc")="1" then%> checked <% end if%>>
        ����&nbsp; <input type="radio" name="fhc" id="fhc" value="0" <% if rs("fhc")="0" then%> checked <% end if%>>
        ������</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���û�ע���ͣ�</strong></td>
      <td height="25" class="tdbg">������
        <input name="songfa" type="text" id="songfa" value="<%=rs("songfa")%>" onKeyUp="if(isNaN(value))execCommand('undo')">
        �������֣�
        <input name="songji" type="text" id="songji" value="<%=rs("songji")%>" onKeyUp="if(isNaN(value))execCommand('undo')">
        ��</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��ʱ�Զ��˳�������ٴβ����ٽ�����</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="exitauto" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("exitauto")%>">
      ��</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>ͬһ����Աֻ�ܽ��ֶ��ٸ�δ��ɵ�����</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="jieweiok" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("jieweiok")%>">
        ��
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>ͬһ����Աֻ�ܷ������ٸ�δ��ɵ�����</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="faweiok" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("faweiok")%>">
      ��</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>һ���˺�ͬʱ����󶨶��ٸ��Ա���ţ�</strong></td>
      <td height="25" class="tdbg"><input type="text" name="buyhao" id="buyhao" size="40" maxlength="250" value="<%=rs("buyhao")%>" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���ٸ����������һ��ֽ�</strong></td>
      <td height="25" class="tdbg"><input name="fhcshu" type="text" id="fhcshu" value="<%=rs("fhcshu")%>" size="40" maxlength="250" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    
    
    <tr>
      <td height="25" class="tdbg"><strong>�Ƿ񵯳�ƽ̨���棺</strong></td>
      <td height="25" class="tdbg">
	  <select name="popup">
        <option value="1" <%if rs("popup")=1 then%> selected="selected" <%end if%>>����</option>
        <option value="2" <%if rs("popup")=2 then%> selected="selected" <%end if%>>�ر�</option>
      </select>      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��վ�������棺</strong></td>
      <td height="25" class="tdbg"><input name="gonggao" type="text" id="gonggao" value="<%=rs("gonggao")%>" size="40" maxlength="250"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��վ����������⣺</strong></td>
      <td height="25" class="tdbg"><input name="kouhao" type="text" id="kouhao" value="<%=rs("kouhao")%>" size="40" maxlength="250"></td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>��վ��ַ��</strong><br>
        ����д����URL��ַ</td>
      <td width="942" height="25" class="tdbg"><input name="SiteUrl" type="text" id="SiteUrl" value="<%=rs("url")%>" size="40" maxlength="255"> 
      (���治Ҫ��/)      </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>LOGO��ַ��</strong><br>
        ����д����URL��ַ</td>
      <td width="942" height="25" class="tdbg"><input name="LogoUrl" type="text" id="LogoUrl" value="<%=rs("logo")%>" size="40" maxlength="255">
          <input type="button" name="Submit" value="�ϴ�ͼƬ" onClick="window.open('../upload_flash.asp?formname=frm1&editname=LogoUrl&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
        (logo��С175*68px) </td>
    </tr>
	 <tr>
      <td height="25" class="tdbg"><strong>��ҳ�ͷ�ͼƬ��</strong></td>
      <td height="25" class="tdbg"><input name="kefu_pic" type="text" id="kefu_pic" value="<%=rs("kefu_pic")%>" size="40" maxlength="255">
      <input type="button" name="Submit2" value="�ϴ�ͼƬ" onClick="window.open('../upload_flash.asp?formname=frm1&editname=kefu_pic&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
      (��С253*130px) </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>��˾��ַ��</strong><br></td>
      <td width="942" height="25" class="tdbg"><input name="BannerUrl" type="text" id="BannerUrl" value="<%=rs("address")%>" size="40" maxlength="255">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�ͷ����ߣ�</strong></td>
      <td height="25" class="tdbg"><input name="phone" type="text" id="phone" value="<%=rs("phone")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��ѯ���ߣ�</strong></td>
      <td height="25" class="tdbg"><input name="tel" type="text" id="tel" value="<%=rs("tel")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>������ѯQQ��</strong></td>
      <td height="25" class="tdbg"><input name="qq5" type="text" id="qq5" value="<%=rs("qq5")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���ֿͷ�QQ��</strong></td>
      <td height="25" class="tdbg"><input name="qq" type="text" id="qq" value="<%=rs("qq")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��ֵ�ͷ�QQ��</strong></td>
      <td height="25" class="tdbg"><input name="qq2" type="text" id="qq2" value="<%=rs("qq2")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���ֿͷ�QQ��</strong></td>
      <td height="25" class="tdbg"><input name="qq3" type="text" id="qq3" value="<%=rs("qq3")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>Ͷ�߿ͷ�QQ��</strong></td>
      <td height="25" class="tdbg"><input name="qq4" type="text" id="qq4" value="<%=rs("qq4")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>������QQ��</strong></td>
      <td width="942" height="25" class="tdbg"><input name="qq6" type="text" id="qq6" value="<%=rs("qq6")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�Ƹ�ͨ�˺ţ�</strong></td>
      <td height="25" class="tdbg"><input name="cft" type="text" id="cft" value="<%=rs("cft")%>" size="40" maxlength="200"></td>
    </tr>
    
    <tr>
      <td height="25" class="tdbg"><strong>���Ƽ��˹���vip���������ٸ������㣺</strong></td>
      <td height="25" class="tdbg"><input name="vip_car_jieducm" type="text" id="vip_car_jieducm" value="<%=rs("vip_car_jieducm")%>" size="40" maxlength="200"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�ӣ�����һ������ɵö��ٻ��֣�</strong></td>
      <td height="25" class="tdbg"><input name="fajifei" type="text" id="fajifei" value="<%=rs("fajifei")%>" size="40" maxlength="200" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
   
    <tr>
      <td height="25" class="tdbg"><strong>���Ƽ���ע��ʱ�Ƽ��˵ö��ٻ��֣�</strong></td>
      <td height="25" class="tdbg"><input name="tjrzhu" type="text" id="tjrzhu" value="<%=rs("tjrzhu")%>" size="40" maxlength="200" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���Ƽ��˽ӣ�����һ�������Ƽ��˵ö��ٻ��֣�</strong></td>
      <td height="25" class="tdbg"><input name="tjrjifei" type="text" id="tjrjifei" value="<%=rs("tjrjifei")%>" size="40" maxlength="200" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>һ���½���ͬһ������������</strong></td>
      <td height="25" class="tdbg"><input name="geshu" type="text" id="geshu" value="<%=rs("geshu")%>" size="40" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
        ��</td>
    </tr>
	<tr>
      <td height="25" class="tdbg"><strong>����һ���ղ�����۶��ٷ����㣺</strong></td>
      <td height="25" class="tdbg"><input name="cang" type="text" id="cang" value="<%=formatnumber(rs("cang"),2,-1)%>" size="40" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���ֻ������������</strong></td>
      <td height="25" class="tdbg"><input name="jifei" type="text" id="jifei" value="<%=rs("jifei")%>" size="40" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
        (����������,���ٸ����ֿ��Ի�һ��������)</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��������ռ۸�</strong></td>
      <td height="25" class="tdbg"><input name="fa" type="text" id="fa" value="<%=formatnumber(rs("fa"),2,-1)%>" size="40" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
        (����������,1�����������Ǯ)</td>
    </tr>
	
    <tr>
      <td height="25" class="tdbg"><strong>��������ۼ۸�</strong></td>
      <td height="25" class="tdbg">��ͨ��Ա�۸�
        <input name="kuan" type="text" id="kuan" value="<%=formatnumber(rs("kuan"),2,-1)%>" size="10" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
        Ԫ
        &nbsp;&nbsp; VIP��Ա�۸�
        <input name="kuanvip" type="text" id="kuanvip" value="<%=formatnumber(rs("kuanvip"),2,-1)%>" size="10" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
       Ԫ       (����������,ÿ�����������Ǯ)</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>������ʱ���ָ��ڣ�</strong>(������÷��������)</td>
      <td height="25" class="tdbg"><input type="text" name="jifeibig" id="jifeibig" value="<%=rs("jifeibig")%>">
        ������ʱ���÷�����
          <input name="kou" type="text" id="kou"  value="<%=formatnumber(rs("kou"),1,-1)%>"  size="20" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
         
(���������ֿ���ΪС����:0.9)����������û��ֵ÷����������1��1</td>
    </tr>
    
     <tr>
      <td height="25" class="tdbg"><strong>�����ֻ������շѱ�׼��</strong></td>
      <td height="25" class="tdbg"><label>
        ��ͨ�û�һ��<input name="puuser" value="<%=rs("puuser")%>" type="text" size="10" maxlength="10" onKeyUp="if(isNaN(value))execCommand('undo')">��
        vip�û�һ��<input name="vipuser" value="<%=rs("vipuser")%>" type="text" size="10" maxlength="10" onKeyUp="if(isNaN(value))execCommand('undo')">��
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>VIP��Ա������ʱ�ö��ٷ����������</strong></td>
      <td height="25" class="tdbg"><input type="text" name="vipjifei" id="vipjifei" value="<%=formatnumber(rs("vipjifei"),1,-1)%>">        (���������ֿ���ΪС����:0.9)</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���ͷ������̷ּѣ�</strong></td>
      <td height="25" class="tdbg"><input name="shouxu" type="text" id="shouxu" value="<%=rs("shouxu")%>" size="40" maxlength="100" onKeyUp="if(isNaN(value))execCommand('undo')">
      &nbsp; (����������,��������10,������ȡ10%���̷ּ�)</td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>վ�����䣺</strong></td>
      <td width="942" height="25" class="tdbg"><input name="WebmasterEmail" type="text" id="WebmasterEmail" value="<%=rs("email")%>" size="40" maxlength="100">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�����̻��ţ�</strong></td>
      <td height="25" class="tdbg"><input name="wu" type="text" id="wu" value="<%=rs("wu")%>" size="40" maxlength="100"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>����MD5��Կ��</strong></td>
      <td height="25" class="tdbg"><input name="shi" type="text" id="shi" value="<%=rs("shi")%>" size="40" maxlength="100"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>֧�����˺ţ�</strong></td>
      <td height="25" class="tdbg"><input name="WebmasterName" type="text" id="WebmasterName" value="<%=rs("zfb")%>" size="40" maxlength="200"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>֧�����˺Ŷ�Ӧ��partnerID��</strong></td>
      <td height="25" class="tdbg"><input name="wushi" type="text" id="wushi" value="<%=rs("wushi")%>" size="40" maxlength="100"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>֧������ȫУ���룺</strong></td>
      <td height="25" class="tdbg"><input name="yibai" type="text" id="yibai" value="<%=rs("yibai")%>" size="40" maxlength="100"></td>
    </tr>
    
    <tr>
      <td height="25" class="tdbg"><strong>�ױ��̻���ţ�</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="jieducm_MerId" size="40" maxlength="100" value="<%=rs("jieducm_MerId")%>">
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�ױ���Կ��</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="jieducm_MerKey" size="40" maxlength="100" value="<%=rs("jieducm_MerKey")%>">
      </label></td>
    </tr>
	<tr>
      <td height="25" class="tdbg"><strong>�Ƿ�ͨ��������۹��ܣ�</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="radio" name="payis" value="1" <%if rs("payis")=1 then%> checked <%end if%>>��ͨ
        <input type="radio" name="payis" value="0" <%if rs("payis")=0 then%> checked <%end if%>>�ر�
      </label></td>
    </tr>
	    <tr>
      <td height="25" class="tdbg"><strong>���۷�������ͻ���Ϊ��</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="payjifei" onKeyUp="if(isNaN(value))execCommand('undo')"  value="<%=rs("payjifei")%>">
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���۷������������Ϊ��</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="paynum" onKeyUp="if(isNaN(value))execCommand('undo')"  value="<%=rs("paynum")%>">
      ��</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���۷�����������Ϊ��</strong></td>
      <td height="25" class="tdbg">VIP��Ա��
        <label>
        <input name="vippaynum" type="text" size="10" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("vippaynum")%>">
     %&nbsp;&nbsp; ��ͨ��Ա��
     <input name="pupaynum" type="text" size="10" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("pupaynum")%>">
      %</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��ѯ��ݵ���ʹ�ô�����</strong></td>
      <td height="25" class="tdbg"><label>
        �������
        <input type="text" name="RndQueryNum" size="10" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("RndQueryNum")%>">
        ��   
       &nbsp; �Զ�������
        <input type="text" name="DefQueryNum" size="10" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("DefQueryNum")%>">
      ��</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>�������棺</strong></td>
      <td height="25" class="tdbg"><label>
        <textarea name="jieducm_gonggao" cols="40" rows="4" id="jieducm_gonggao"><%=rs("jieducm_gonggao")%></textarea>
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>����ͳ�ƴ��룺<br>
      </strong>Ҳ�ɼ���QQ��ѯ̨����</td>
      <td height="25" class="tdbg">
	  <textarea name="icp" cols="40" rows="4" id="icp"><%=rs("icp")%></textarea>
      &nbsp;</td>
    </tr>

     <tr> 
      <td height="25" colspan="2" class="title"><a name="Email"></a><strong>�ʼ�������ѡ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong><a href="#Top"><font color="#0066CC">����&nbsp;����</font></a></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>�ʼ����������</strong><br>
        ��һ��Ҫѡ����������Ѱ�װ�����</td>
      <td height="25" class="tdbg"> 
        (�������� JMAIL ֧��)</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP��������ַ��</strong><br>
        ���������ʼ���SMTP������<br>
        ����㲻����˲������壬����ϵ��Ŀռ��� </td>
      <td height="25" class="tdbg"> 
        <input name="MailServer" type="text" id="MailServer" value="<%=rs("MailServer")%>" size="40">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP��¼�û�����</strong><br>
        ����ķ�������ҪSMTP�����֤ʱ�������ô˲���</td>
      <td height="25" class="tdbg"> 
        <input name="MailServerUserName" type="text" id="MailServerUserName" value="<%=rs("MailServerUserName")%>" size="40">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP��¼���룺</strong><br>
        ����ķ�������ҪSMTP�����֤ʱ�������ô˲��� </td>
      <td height="25" class="tdbg"> 
        <input name="MailServerPassWord" type="PassWord" id="MailServerPassWord" value="<%=rs("MailServerPassWord")%>" size="40">      </td>
    </tr>
    
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP����</strong>��<br>
        ����á�xieshanhui@163.com���������û�����¼ʱ����ָ��778892.com</td>
      <td height="25" class="tdbg"> 
        <input name="MailDomain" type="text" id="MailDomain" value="<%=rs("MailDomain")%>" size="40">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���û�ע���Ƿ�E_mail��</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="checkbox" name="isemail" value="1" <%if rs("isemail")=1 then%> checked="checked" <%end if%>>
      ѡ�����û�ע��󽫷��ͻ�ӭ�ʼ�</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>���û�ע���Ƿ�վ����Ϣ��</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="checkbox" name="ismessage" value="1" <%if rs("ismessage")=1 then%> checked="checked" <%end if%>>
      ѡ�������û�ע�ᷢ��վ����Ϣ</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>Email�������ã�</strong></td>
      <td height="25" class="tdbg"><label>
        <textarea name="emailcontent" cols='90' rows='6'><%=rs("emailcontent")%></textarea>
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>վ����Ϣ���ݣ�</strong></td>
      <td height="25" class="tdbg"><textarea name="messagecontent" cols='90' rows='6'><%=rs("messagecontent")%></textarea></td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveConfig">
          <input name="cmdSave" type="submit" id="cmdSave" value=" &nbsp;��������&nbsp; " <% 'If ObjInstalled=false Then response.write "disabled" %> style="cursor: hand;background-color: #cccccc;">      </td>
    </tr>
</td>
      </tr>
  </form>
  </table>
<%end if%>
  <!--#include file="Admin_fooder.asp" -->
</body>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</html>
