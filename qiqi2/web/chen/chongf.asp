 <!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../md5.asp"-->
 <%
 '------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
 id=request("id")
if request.Form<>"" then
 username=request("username")
 cunkuan=request("cunkuan")
 fabudian=request("fabudian")
 jifei=request("jifei")
 class1=request("class")
 content=request("content")
 czm=request("czm")

IF not isNumeric(request("cunkuan")) then
   Response.Write("<script>alert('����������Ϊ���֣����������룡');history.back();</script>")
   response.End()
elseif not isNumeric(request("fabudian")) then
   Response.Write("<script>alert('�������������Ϊ���֣����������룡');history.back();</script>")
   response.End()
elseif not isNumeric(request("jifei")) then
   Response.Write("<script>alert('�������ֱ���Ϊ���֣����������룡');history.back();</script>")
   response.End()
end if


if md5(czm)<>admin_czmpass then
   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")=class1
   rs("content")=session("AdminName")&"����Ա��"&username&"���к�̨��ֵ�������������"
   rs("jiegou")="ʧ��"
   rs.update
   rs.close
   
     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")=class1
		rs("content")=session("AdminName")&"����Ա��"&username&"���к�̨��ֵ�������������"
        rs("price")=0	
		rs("jiegou")="ʧ��"
		rs("now")=date()
    	rs.update
    	rs.close
   Response.Write("<script>alert('����������!�벻Ҫ�ظ�������ƽ̨���¼�����Ϊ��');history.back();</script>")
   response.End()
end if

if cunkuan="" then
 cunkuan=0
end if
if fabudian="" then
 fabudian=0
end if
if jifei="" then
 jifei=0
end if 



 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select cunkuan,fabudian,jifei From "&jieducm&"_register where username='"&username&"'" ,Conn,3,3  
		rs("cunkuan")=rs("cunkuan")+cunkuan
		rs("fabudian")=rs("fabudian")+fabudian
		rs("jifei")=rs("jifei")+jifei
    	rs.update
    	rs.close
  
if cunkuan<>0 then
  dian=cunkuan&"Ԫ���"
end if
if fabudian<>0 then
  dian=dian&fabudian&"��������"
end if
if jifei<>0 then
  dian=dian&jifei&"������"
end if

  	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")=class1
		rs("content")=session("AdminName")&"����Ա��"&username&"����"&dian
        rs("price")=cunkuan
		rs("jiegou")="�ɹ�"
		rs("now")=date()
    	rs.update
    	rs.close
	
	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu" ,Conn,3,3  
	    rs.addnew
		rs("class")="��̨��ֵ"
		rs("username")=username
    	rs("fabudian")=fabudian
		rs("cunkuan")=cunkuan
		rs("jifei")=jifei
		rs("now")=now()
    	rs.update
    	rs.close  
		
	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")=class1
		rs("content")=session("AdminName")&"����Ա��"&username&"����"&dian
		rs("jiegou")="�ɹ�"		
    	rs.update
    	rs.close
   Response.Write("<script>alert('��ֵ�ɹ�!');window.location.href='usermanage.asp?action=sear&text="&username&"';</script>")
end if
 Set rs=server.createobject("ADODB.RECORDSET")
 rs.open "Select * From "&jieducm&"_register where id="&id&"",Conn,3,3
 if not rs.eof then
%>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<SCRIPT language=javascript>
function  save_onclick22()
{

if   (!(document.aspnetForm.fabudian.value.indexOf(".")   ==   -1)){   
  alert("���������㲻��ΪС��");   
  document.aspnetForm.fabudian.focus();   
  return   false;   
  }	
if   (!(document.aspnetForm.cang.value.indexOf(".")   ==   -1)){   
  alert("�����ղص㲻��ΪС��");   
  document.aspnetForm.cang.focus();   
  return   false;   
  }	
  if   (!(document.aspnetForm.jifei.value.indexOf(".")   ==   -1)){   
  alert("�������ֲ���ΪС��");   
  document.aspnetForm.jifei.focus();   
  return   false;   
  }	
  return true;  
}
 
  function   test(str){   
          var   pos;   
          var   fst   
          var   lst;   
          if   (str   ==   "")   return;   
          pos   =   str.indexOf(".");   
          if   (pos   !=   -1){   
                  fst   =   str.substring(0,pos);   
                  lst   =   str.substring(pos+1,pos.length);   
                  if   (lst.length   >   1){                             
                            document.aspnetForm.cunkuan.value="";   
							  alert("������������λС��");   

                  }   
          }           
  }  
  function   test1(str){   
          var   pos;   
          var   fst   
          var   lst;   
          if   (str   ==   "")   return;   
          pos   =   str.indexOf(".");   
          if   (pos   !=   -1){   
                  fst   =   str.substring(0,pos);   
                  lst   =   str.substring(pos+1,pos.length);   
                  if   (lst.length   >   1){                             
                            document.aspnetForm.fabudian.value="";   
							  alert("���������㲻������λС��");   

                  }   
          }           
  }  
  function   test2(str){   
          var   pos;   
          var   fst   
          var   lst;   
          if   (str   ==   "")   return;   
          pos   =   str.indexOf(".");   
          if   (pos   !=   -1){   
                  fst   =   str.substring(0,pos);   
                  lst   =   str.substring(pos+1,pos.length);   
                  if   (lst.length   >   1){                             
                            document.aspnetForm.cang.value="";   
							  alert("�����ղص㲻������λС��");   

                  }   
          }           
  }  
  function   test3(str){   
          var   pos;   
          var   fst   
          var   lst;   
          if   (str   ==   "")   return;   
          pos   =   str.indexOf(".");   
          if   (pos   !=   -1){   
                  fst   =   str.substring(0,pos);   
                  lst   =   str.substring(pos+1,pos.length);   
                  if   (lst.length   >   1){                             
                            document.aspnetForm.jifei.value="";   
							  alert("�������ֲ�������λС��");   

                  }   
          }           
  }   
  </script>   
    
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 10px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<FORM id=aspnetForm name=aspnetForm  action="" method=post  onSubmit="return save_onclick22();">
<DIV> </DIV>
<DIV style="WIDTH: 98%; PADDING-TOP: 10px"><SPAN 
id=ctl00_AllBaseContentPlaceHolder_MsgLabel style="COLOR: red"></SPAN><SPAN 
id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator1 
style="DISPLAY: none; COLOR: red">* ��ֵ����Ϊ�գ�</SPAN> <SPAN 
id=ctl00_AllBaseContentPlaceHolder_RangeValidator1 
style="DISPLAY: none; COLOR: red">* ��ֵ��Χ��0.01-999֮�䣡</SPAN> <SPAN 
id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator2 
style="DISPLAY: none; COLOR: red">* ��ֵ�û�������Ϊ�գ�</SPAN></DIV>
<DIV style="WIDTH: 95%; PADDING-TOP: 10px; HEIGHT: 160px">
<DIV style="PADDING-LEFT: 50px; WIDTH: 98%; HEIGHT: 20px">
  <div align="center"><SPAN 
style="FONT-WEIGHT: bolder; FONT-SIZE: 14px; COLOR: red">���߳�ֵ����������</SPAN><BR>
  </div>
</DIV>
<UL class=pushmission>
  <LI>
    <table width="100%" border="0">
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td width="39%"><div align="right">��ֵ�û�����</div></td>
        <td width="61%"><input type="text" name="name" id="name"  value="<%=rs("username")%>" readonly>
          ��<%
if rs("cunkuan")=0 then
response.Write("0.0")
else
response.Write(formatnumber(rs("cunkuan"),1))
end if
%> �����㣺<%
if rs("fabudian")=0 then
response.Write("0.0")
else
response.Write(formatnumber(rs("fabudian"),1))
end if
%>            ���֣�<%=rs("jifei")%></td>
      </tr>
      <tr>
        <td><div align="right">������</div></td>
        <td>
        <INPUT id=cunkuan    name=cunkuan maxLength=4    value="0"  onkeyup="test(this.value)"   />
              
          <span class="STYLE1">X��-X ����������</span></td>
      </tr>
      <tr>
        <td><div align="right">���������㣺</div></td>
        <td><input name="fabudian" type="text" id="fabudian"   maxLength=4   value="0"  onkeyup="test1(this.value)"/>
          <span class="STYLE1">X��-X ����������</span></td>
      </tr>
      
    
      <tr>
        <td><div align="right">�������֣�</div></td>
        <td><input name="jifei" type="text" id="jifei"   maxLength=5  value="0"  onkeyup="test3(this.value)"/>
          <span class="STYLE1">X��-X ����������</span></td>
      </tr> 
      
       <tr>
        <td><div align="right">�������ͣ�</div></td>
        <td><input name="class" type="radio" id="class" value="����" checked="checked" />
          ���� <span class="STYLE1">*�ȱ�ˢ����ʱ�������Ҫ�޸�</span></td>
      </tr> 
       <tr>
         <td>&nbsp;</td>
         <td><input type="radio" name="class" id="class" value="��ֵ" />
           ��ֵ <span class="STYLE1">*�û���ʵ��ֵ�������߳�ֵʧ�ܣ�֧������ֵ��</span></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td><input type="radio" name="class" id="class" value="��ֵ" />
           ��ֵ <span class="STYLE1">*�û���ʵ��ֵ���繺�򷢲��㣬�һ������㣬����</span></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td><input type="radio" name="class" id="class" value="����" />
           ���� <span class="STYLE1">*�������á�</span></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td><input type="radio" name="class" id="class" value="����" />
           ���� <span class="STYLE1">*�����Լ��˳�ֵ���ͷ�������ָ��û���</span></td>
       </tr>
       
        <tr>
         <td><div align="right">�����룺</div></td>
         <td><label>
           <input type="password" name="czm" />
         </label></td>
       </tr>
       <tr>
         <td><div align="right">����:</div></td>
         <td><textarea name="content" id="content" cols="45" rows="5"></textarea></td>
       </tr>
      <tr>
        <td colspan="2"><div align="center">
         <input type="hidden" name="username" value="<%=rs("username")%>" />
		  <input type="submit" name="button" id="button" value="�� ��">
          &nbsp; 
          <input type="button" name="button2" id="button2" value="ȡ ��"  onClick="history.back();"/>
        </div></td>
        </tr>
    </table>
  </LI>
</UL>
</DIV>


<DIV></DIV>

</FORM></DIV>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
<% end if%>