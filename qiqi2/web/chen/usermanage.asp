<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�� �� �� ��</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>

<body>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
    <td colspan="2" align="center" class="title"><strong>�� �� �� ��</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30" ><strong>��������</strong></td>
    <td><a href="usermanage.asp">���л�Ա </a>|<a href="usermanage.asp?action=pu">δ��˻�Ա </a>| <a href="usermanage.asp?action=vip">����˻�Ա</a>| <a href="usermanage.asp?action=taobao">�Ա�����Ա </a>|<a href="usermanage.asp?action=paipai">��������Ա </a>| <a href="usermanage.asp?action=youa">�а�����Ա </a>|<a href="usermanage.asp?action=you">����� |</a> <a href="usermanage.asp?action=wu">�����</a> | <a href="?action=isjie">��ֹ������</a> | <a href="?action=isfa">��ֹ������</a>| <a href="?action=isdun">��ֹ�һ�</a> | <a href="?action=islogin">��ֹ��¼</a> | <a href="?action=isgive">��ֹ����</a> | <a href="?action=isgives">���Ͳ���������</a> | <a href="?action=ismessage">��ֹվ����</a>| <a href="?action=wangwang">�Ѱ��ֻ�</a> | <a href="?action=vipuser">VIP��Ա</a> | <a href="?action=lastnow">���µ�½</a>|<a href="?action=zuan">��½���</a></td>
  </tr>
</table>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="usermanage.asp?action=sear">

                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>�����û�����
                  <input class=input1 size=25 name=text>
                  
                  <input type="checkbox" name="hao" id="hao" value="1">
                  ���Ա�С�������� 
                  <input type="checkbox" name="da" id="da" value="1">��ѯ���Ա����
				   <input type="checkbox" name="phone" id="phone" value="1">
				     ��ѯ���ֻ���
				     <input type="checkbox" name="qq" id="qq" value="1">��ѯqq��
<input name="submit" type=submit class=input2 value=" �� �� ">
                </form> 
            </div></td>
  </tr>
          
          </td>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center><tr></tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0" ><tr>
   
      <table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
        <tr>
          <td colspan="2" align="left" class="title"></td>
        </tr>
        <tr valign="middle" class="tdbg">
          <td height="22"></td>
          <td width="200" height="22" align="right"></td>
        </tr>
      </table>
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22">
            <td width="131" height="22" align="center" ><strong>�û���</strong></td>
            <td width="115" align="center" ><strong>��ϵ</strong></td>
            <td width="297" align="center" ><strong>��� ������  ����   ע��ʱ��</strong></td>
            <td width="159" align="center" ><strong>��������</strong></td>
            <td width="157" align="center" ><strong>��������</strong></td>
            <td width="270" align="center" ><strong>�����</strong></td>
            <td width="248" align="center" ><strong>����</strong></td>
            <!--<td width="40" align="center" ><strong>�����</strong></td>-->
          </tr>
       
			<%	
			action=request("action")
			hao=request("hao")
			da=request("da")
			phone=request("phone")
			qq=Replace(Trim(Request.Form("qq")),"'","''")
			if action="you" then
			sql="SELECT * FROM "&jieducm&"_register where cunkuan<>0  order by cunkuan desc"
			elseif action="wu" then
			sql="SELECT * FROM "&jieducm&"_register where cunkuan=0 order by id desc"
			elseif action="sear"  and hao<>"1"  and da<>"1" and phone <>"1"  and qq<>"1" then
			   	sql="SELECT * FROM "&jieducm&"_register where username like '%"&request("text")&"%'  order by id desc"
			elseif action="sear" and qq="1" then	
			   sql="SELECT * FROM "&jieducm&"_register where qq='"&request("text")&"' order by id desc"	
			elseif  action="sear" and phone="1" then
			   Sql = "select * from "&jieducm&"_register where dst="&request("text")&""	
			elseif action="pu" then
			   sql="SELECT * FROM "&jieducm&"_register where level1='0' order by id desc"
			elseif action="vip" then
				sql="SELECT * FROM "&jieducm&"_register where level1='1' order by id desc"
			elseif action="isjie" then
				sql="SELECT * FROM "&jieducm&"_register where isjie=1 order by id desc"
			elseif action="isfa" then
				sql="SELECT * FROM "&jieducm&"_register where isfa=1 order by id desc"
			elseif action="isdun" then
				sql="SELECT * FROM "&jieducm&"_register where isdun=1 order by id desc"
			elseif action="islogin" then
				sql="SELECT * FROM "&jieducm&"_register where islogin=1 order by id desc"
			elseif action="isgive" then
				sql="SELECT * FROM "&jieducm&"_register where isgive=1 order by id desc"
			elseif action="isgives" then
				sql="SELECT * FROM "&jieducm&"_register where isgives=1 order by id desc"
			elseif action="ismessage" then
				sql="SELECT * FROM "&jieducm&"_register where ismessage=1 order by id desc"
			elseif action="wangwang" then
				sql="SELECT * FROM "&jieducm&"_register where dst<>0 order by id desc"
			elseif action="taobao" then
				sql="SELECT * FROM "&jieducm&"_register where taobao=1 order by id desc"
			elseif action="paipai" then
				sql="SELECT * FROM "&jieducm&"_register where paipai=1 order by id desc"
			elseif action="youa" then
				sql="SELECT * FROM "&jieducm&"_register where youa=1 order by id desc"	
			elseif action="vipuser" then
				sql="SELECT * FROM "&jieducm&"_register where vip=1 order by id desc"		
			elseif action="lastnow" then
				sql="SELECT * FROM "&jieducm&"_register  order by lastnow desc"	
			elseif action="zuan" then
					sql="SELECT * FROM "&jieducm&"_register  order by zuan desc"					
			elseif action="sear" and hao="1" then
			    Sql = "select * from "&jieducm&"_register where username=(Select username From "&jieducm&"_xinyu where shopname='"&request("text")&"' and class=1) order by id desc"
		    elseif  action="sear" and da="1" then
			   Sql = "select * from "&jieducm&"_register where username=(Select username From "&jieducm&"_keeper where keeper='"&request("text")&"')  order by id desc"
			else
			   	sql="SELECT * FROM "&jieducm&"_register  order by level1 ,id desc"
			end if
			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
	dim maxperpage,url,PageNo
	if action="" then
	 url="usermanage.asp"
	 elseif action="sear"  then
	url="usermanage.asp?action=sear&text="&request("text")&""
	else
	 url="usermanage.asp?action="&action&""
	 end if
	 rs.pagesize=10
	 PageNo=REQUEST("PageNo")
	 if PageNo="" or PageNo=0 then PageNo=1
	 RS.AbsolutePage=PageNo
	 TSum=rs.pagecount
	 maxperpage=rs.pagesize
	 RowCount=rs.PageSize
	   PageNo=PageNo+1
	   PageNo=PageNo-1
	 if CINT(PageNo)>1 then
	    if CINT(PageNo)>CINT(TSum) then
		  response.Write("�Բ���û������Ҫ��ҳ��")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("û����һҳ!")
		Response.End
	 End if
	i=0
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	 DO WHILE NOT rs.EOF AND RowCount>0
	 username=rs("username")
	 qq=rs("qq")
	 registerip=rs("registerip")
	 LastLoginIP=rs("LastLoginIP")
	 %>
        <form name="myform<%=i%>" method="Post" action="useredit2.asp" >
		  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
		    <td width="131" align="center">
             
            <a href="admin_recordu.asp?username=<%=username%>" title="�鿴�˻�Ա�Ĳ�����¼"> <%=username%></a><%if Replace(Trim(rs("vip")),"'","''")	="1" then%><img src="../images/jieducm_vip.gif"  title="vip��Ա" /><%end if%><br>�Ƽ��ˣ�<%=rs("tjr")%><br>
			<%if rs("level1")="0" then %><font color="#FF0000"><strong>δ���</strong></font><%else%><font color="#0000FF"><strong>�����</strong></font><%end if%><br>
			<%if rs("activestart")="0" then %><font color="#FF0000"><strong>δ�ʼ�����</strong></font>
			<%elseif  rs("activestart")="1" then%><font color="#0000FF"><strong>���ʼ�����</strong></font><%end if%>
			<br>
				<%if rs("dst")="0" then %><font color="#FF0000"><strong>δ���ֻ�</strong></font>
			<%else%><font color="#0000FF"><strong>�Ѱ��ֻ�<%=rs("dst")%></strong></font><%end if%>
			<br><%if rs("tribeid")=0 then%><font color="#FF0000"><strong>δ���벿��</strong></font>
			<%else%><font color="#0000FF"><strong>�Ѽ��벿��</strong></font><%end if%>
			</td>
            <td width="115" align="center">
			<%Sql1 = "select  count(*) as jieducm_tjr  from "&jieducm&"_register where tjr='"&username&"' "
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
				Rs1.Open Sql1,conn,1,1
				if NOT rs1.EOF  then
				jieducm_tjr=rs1("jieducm_tjr")
            end if%>�Ƽ�������<font color="#FF0000"><strong><%=jieducm_tjr%></strong></font>��
			 <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes">
			<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0"  title="QQ��<%=qq%>" /></a></td>
            <td width="297" align="center"><div align="left"><a href="chongf.asp?id=<%=rs("id")%>" title="�Դ˻�Ա��̨��ֵ"><font color="#FF0000"><strong>��
                <%
if rs("cunkuan")=0 then
response.Write("0.0")
else
response.Write(formatnumber(rs("cunkuan"),2))
end if
%> 
                �����㣺
                <%
if rs("fabudian")=0 then
response.Write("0.0")
else
response.Write(formatnumber(rs("fabudian"),2))
end if
%>
  ���֣�<%=rs("jifei")%></strong></font></a><br>
  ע��ʱ�䣺<%=rs("now")%><br>
  ע��ip��<a href="http://www.ip138.com/ips.asp?ip=<%=registerip%>&action=2" target="_blank" title="�鿴ip���ڵ�"><%=registerip%></a>
  &nbsp;<a href="jieducm_ip.asp?ip=<%=registerip%>"><font color="#FF0000">������IP</font></a><br>
  ����½ʱ�䣺<%if isnull(RS("lastnow")) then%><%=rs("now")%><%else%><%=rs("lastnow")%><%end if%><BR> 
  ����½ip��
  <%if not isnull(rs("LastLoginIP")) then%>
  <a href="http://www.ip138.com/ips.asp?ip=<%=LastLoginIP%>&action=2" target="_blank" title="�鿴ip���ڵ�"><%=LastLoginIP%></a>
  &nbsp;<a href="jieducm_ip.asp?ip=<%=LastLoginIP%>"><font color="#FF0000">������IP</font></a>
  <%else%>
  <a href="http://www.ip138.com/ips.asp?ip=<%=registerip%>&action=2" target="_blank" title="�鿴ip���ڵ�"><%=registerip%></a>
  &nbsp;<a href="jieducm_ip.asp?ip=<%=registerip%>"><font color="#FF0000">������IP</font></a>
  <%end if%>
  
  <br>
  ��½������<%=rs("zuan")%>��
  <br>
  <a href="jieducm_usermoney.asp?username=<%=username%>"><Font color="#FF0000"><strong>�鿴�˻�Ա�ʽ�����</strong></Font></a>
  </div></td>        
            <td width="159" align="center"> 
           <% Sql1 = "select  count(*) as renwu  from "&jieducm&"_zhongxin where username='"&username&"' "
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
				Rs1.Open Sql1,conn,1,1
				if NOT rs1.EOF  then
				renwu=rs1("renwu")
            end if
			
			totalprices=0
			Sql1 = "select  sum(price) as totalprices  from "&jieducm&"_zhongxin where username='"&username&"'  and classid<>'6'"
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
			Rs1.Open Sql1,conn,1,1
			if NOT rs1.EOF  then
				totalprices=rs1("totalprices")
		    else
				totalprices=0
            end if
		    rs1.close
			  %>  
			  <%=renwu%>��<br>
			  
			  <a href="admin_mymission.asp?sort=ok&text=<%=username%>">�鿴�ѷ�����</a><br><br>
			  �ѷ������ܶ<font color="#FF0000"><%if isnull(totalprices) then%> 0 <%else%><%=totalprices%><%end if%></font>Ԫ
            </td> 
            <td width="157" align="center">     
		  <% Sql2 = "select  count(*) as jiewu  from "&jieducm&"_jieshou where username='"&username&"' "
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
				Rs1.Open Sql2,conn,1,1
				if NOT rs1.EOF  then
				jiewu=rs1("jiewu")
				end if
				
			Sql1 = "select  sum(price) as totalpricesj  from "&jieducm&"_jieshou where username='"&username&"'  and classid<>'6'"
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
			Rs1.Open Sql1,conn,1,1
			if NOT rs1.EOF  then
				totalpricesj=rs1("totalpricesj")
				else
				totalpricesj=0
            end if
				rs1.close  %> 
				  <%=jiewu%>��<br>
				    <a href="admin_MyMissionjie.asp?sort=ok&text=<%=username%>">�鿴�ѽ�����</a><br><br>
				  �ѽ������ܶ<font color="#FF0000"><%if isnull(totalpricesj) then%> 0 <%else%><%=totalpricesj%><%end if%></font>Ԫ
		    </td>         
            <td width="270" align="center">
               <label> <input type="checkbox" name="level1" id="level1"  value="1" <%if rs("level1")="1" then%> checked <% end if%> > ��˻�Ա </label>
              <label> <input type="checkbox" name="taobao" id="taobao"  value="1" <%if rs("taobao")="1" then%> checked <% end if%> > �Ա�����Ա<br>
</label> <label> <input type="checkbox" name="paipai" id="paipai"  value="1" <%if rs("paipai")="1" then%> checked <% end if%> > ��������Ա </label>

               <label> <input type="checkbox" name="isjie" id="isjie"  value="1" <%if rs("isjie")=1 then%> checked <% end if%> > ��ֹ������ </label><br>
               <label><input type="checkbox" name="isfa" id="isfa" value="1" <% if rs("isfa")=1 then%> checked <% end if%>>��ֹ������</label>
            <label> <input type="checkbox" name="isdun" id="isdun" value="1" <%if rs("isdun")=1 then%> checked <%end if%> >��ֹ�һ�</label><br>
             <label><input type="checkbox" name="islogin" id="islogin" value="1" <%if rs("islogin")=1 then%> checked <%end if%>>��ֹ��¼</label>
             
           
              <label> <input type="checkbox" name="isgive" id="isgive" value="1" <%if rs("isgive")=1 then%> checked <% end if%>>��ֹ����</label><br>
             <label><input type="checkbox" name="isgives" id="isgives" value="1" <% if rs("isgives")=1 then%> checked <% end if%>>���Ͳ����̷ּ�</label>
              <label><input name="ismessage" type="checkbox" id="ismessage" value="1" <% if rs("ismessage")=1 then%> checked <% end if%>>��ֹվ����</label>
 
              <input type="hidden" name="id" id="id" value="<%=rs("id")%>">
              <input type="hidden" name="username1" id="username1" value="<%=rs("username")%>">
            <br>            </td>
          
            <td width="248" align="center" >
			<a href="userunionlist.asp?username=<%=rs("username")%>">�鿴�˻�Ա���ƹ��¼</a><br>
			<a onClick="javascript:document.forms['myform<%=i%>'].submit();" style="cursor:pointer;">�޸������</a>|<a href="chengfa.asp?username=<%=username%>">�����ͷ�</a>
              <br>
             <a href="useredit.asp?username=<%=username%>"> �޸�����</a>| <a onClick="return confirm('ȷ��Ҫɾ���˻�Ա�𣿴˲������ָܻ���');" href="userdel.asp?id=<%=rs("id")%>">ɾ��</a><br>
            
          <a href="userbuyhao2.asp?id=<%=rs("id")%>">�����</a>|<a href="userbuyhao.asp?username=<%=username%>">�������</a>
		  <br>
		  <a href="usermyww.asp?id=<%=rs("id")%>">���ƹ�</a>| <a href="userbuyww.asp?username=<%=username%>">�����ƹ�</a><br>
		  <%if rs("codenum")<>0 then%><a href="uphonelook.asp?id=<%=rs("id")%>">�鿴�ֻ���֤��</a><br><%end if%>
		   <a href="jieducm_vipuser.asp?username=<%=username%>">����Ϊvip��Ա</a>
		  </td>   
          </tr></form>
 <%
	i=i+1
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          </table>
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg">
    <td height="30" colspan="2">
&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr bgcolor="#FFFFFF"><td  ><div align="center"><br>
    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %>
  </div></td></tr>
</table>
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>���ߴ�ý-���߻�ˢƽ̨</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>

<%rs.close
set rs=nothing
set rs1=nothing
conn.close
set conn=nothing%>
</body>
</html>
