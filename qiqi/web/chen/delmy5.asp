<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
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
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
id=request("id")
action=request("action")
set Rs=server.createobject("adodb.recordset")
sql="select fajifei,kou,vipjifei,cang from "&jieducm&"_stup "
Rs.open sql,conn,1,1
if not(Rs.eof)  then
stupfajifei=rs("fajifei")
stupkou=rs("kou")
stupvipjifei=rs("vipjifei")
stupcang=rs("cang")
end if
if  action="qing" then
  
  	     
elseif action="del" then

      Sqli = "select * from "&jieducm&"_zhongxin where id="&id&" and jieshou_num=0"
	  Set rs = Server.CreateObject("Adodb.RecordSet")
	  rs.Open Sqli,conn,1,1
	  if not rs.eof or not rs.bof then				
       username=rs("username")
       num=rs("num")
	  else
	  	Response.Write("<script>alert('���������ڽ����У�����ɾ��');history.back();</script>")
        response.End()	
      end if

fabu=num*stupcang
sqlr="update "&jieducm&"_register set fabudian=fabudian+"&fabu&"  where username='"&username&"'"
conn.execute sqlr

       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "delete  from "&jieducm&"_zhongxin where id in ("&id&")",conn,3,3
	   	   
	   
	  num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")="����ɾ��"
		rs("content")=session("AdminName")&"��������ɾ��"&id&"����,���������㣺"&fabu&"��"
		rs("price")=0
		rs("jiegou")="����ɾ���ɹ�"
    	rs.update
    	rs.close
		
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")="����ɾ��"
		rs("content")=session("AdminName")&"��������ɾ��"&id&"����"
		rs("price")=0
		rs("jiegou")="����ɾ���ɹ�"
    	rs.update
    	rs.close
			   
set rs=nothing
conn.close
set conn=nothing		
response.Redirect("admin_mymission5.asp")  	   
elseif action="qingli" then

       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "select * from "&jieducm&"_jieshou where id='"&id&"' and start='1'",conn,1,1
       if not (rs.eof) then
	       usrnamef=rs("username")
		   pid=rs("pid")
		else
	      Response.Write("<script>alert('��������!�������Ѳ����ڻ������');history.back();</script>")
          response.End()
		end if
		
      sqlr="delete  from "&jieducm&"_jieshou where id = ('"&id&"')"
      conn.execute sqlr
	
	   Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "select * from "&jieducm&"_zhongxin where pid='"&pid&"'",conn,3,3
       if not (rs.eof) then
           rs("jieshou_num")=rs("jieshou_num")-1
           rs.update
		   rs.close
		else
	      Response.Write("<script>alert('��������!');history.back();</script>")
          response.End()
		end if
		
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=usrnamef
    	rs("num")=num
		rs("class")="�������"
		rs("content")=session("AdminName")&"�����������"&pid&"����"
		rs("price")=0
		rs("jiegou")="������ҳɹ�"
    	rs.update
    	rs.close
		
	   Set rs=server.createobject("ADODB.RECORDSET")
	   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")="�������"
		rs("content")=session("AdminName")&"�����������"&pid&"����"
		rs("price")=0
		rs("jiegou")="������ҳɹ�"
    	rs.update
    	rs.close
		
set rs=nothing
conn.close
set conn=nothing		
response.Redirect("admin_MyMissionjie5.asp")
elseif action="ok" then

		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where id='"&id&"' and start='1'",conn,3,3
           if not (rs.eof) then
		      jusername=rs("username")
			  pid=rs("pid")
		      rs("start")=2
              rs.update
			  rs.close
			else
			  Response.Write("<script>alert('��������!�������Ѳ����ڻ������');history.back();</script>")
               response.End()
			end if
			
       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "select * from "&jieducm&"_zhongxin where pid='"&pid&"'",conn,3,3
       if not (rs.eof) then
	       fusername=rs("username")
		end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_register where username='"&jusername&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+stupfajifei
rs("fabudian")=rs("fabudian")+stupcang
rs.update
rs.close
	   	
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=jusername
    	rs("num")=num
		rs("class")="��̨����ղ�"
		rs("content")=session("AdminName")&"���к�̨�����������"&id&"�������ѵõ�"&stupcang&"��������"&stupfajifei&"����"
		rs("price")=0
		rs("jiegou")="��̨����ղ�"
    	rs.update
    	rs.close 
			   
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=fusername
    	rs("num")=num
		rs("class")="��̨����ղ�"
		rs("content")=session("AdminName")&"���к�̨�����������"&id&"���񣬽��ַ��ѵõ�"&stupcang&"��������"&stupfajifei&"����"
		rs("price")=0
		rs("jiegou")="��̨����ղ�"
    	rs.update
    	rs.close 
		
set rs=nothing
conn.close
set conn=nothing		
response.Redirect("admin_MyMissionjie5.asp") 
end if

%>