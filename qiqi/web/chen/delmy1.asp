<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="../config.asp"-->
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
action=request("action")

set Rs=server.createobject("adodb.recordset")
sql="select fajifei,kou,vipjifei from "&jieducm&"_stup "
Rs.open sql,conn,1,1
if not(Rs.eof)  then
stupfajifei=rs("fajifei")
stupkou=rs("kou")
stupvipjifei=rs("vipjifei")
end if

if  action="qing" then

       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and start='1'",conn,3,3
       if not (rs.eof) then
           rs("start")="0" 
           rs.update
		   rs.close
		else
		 Response.Write("<script>alert('��������!����״̬�ѷ����仯��');history.back();</script>")
          response.End()
	   end if
		   
     sqlr="delete  from "&jieducm&"_jieshou where pid = '"&id&"'"
     conn.execute sqlr
		   
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="�������"
		rs("content")=session("AdminName")&"�����������"&id&"����"
		rs("price")=renum
		rs("jiegou")="������ҳɹ�"
    	rs.update
    	rs.close
        response.Redirect("admin_mymission1.asp")

elseif action="del" then
        
		        Sqli = "select * from "&jieducm&"_zhongxin where pid='"&id&"'  and start='0'"
				Set rs = Server.CreateObject("Adodb.RecordSet")
				rs.Open Sqli,conn,1,1
		        if not rs.eof  then				
                 price=rs("price")
                 username=rs("username")
				 jieducm_fb=rs("jieducm_fb")
				 else
				 Response.Write("<script>alert('��������!����״̬�ѷ����仯��');history.back();</script>")
                response.End()
                end if
				

sqlr="update "&jieducm&"_register set fabudian=fabudian+"&jieducm_fb&",cunkuan=cunkuan+"&price&" where username='"&username&"'"
conn.execute sqlr

       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "delete  from "&jieducm&"_zhongxin where pid in ('"&id&"')",conn,3,3
  
	  num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="����ɾ��"
		rs("content")=session("AdminName")&"��������ɾ��"&id&"����"
		rs("price")=renum
		rs("jiegou")="����ɾ���ɹ�"
    	rs.update
    	rs.close
       response.Redirect("admin_mymission1.asp")
	   
elseif action="back" then
      Set rs=server.createobject("ADODB.RECORDSET")
      rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
       if not (rs.eof) then
           rs("start")=int(rs("start"))-1
           rs.update
		   rs.close
		end if
		   
		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"'",conn,3,3
           if not (rs.eof) then
		   rs("start")=int(rs("start"))-1
           rs.update
		   rs.close
		   end if
		   
	num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="�����ϼ�"
		rs("content")=session("AdminName")&"���з����ϼ�"&id&"����"
		rs("price")=renum
		rs("jiegou")="�����ϼ��ɹ�"
    	rs.update
    	rs.close 
        response.Redirect("admin_mymission1.asp")

elseif action="ok" then

      Set rs=server.createobject("ADODB.RECORDSET")
      rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
       if not (rs.eof) then
           rs("start")=3
           rs.update
		   rs.close
	    end if
		   
		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"'",conn,3,3
           if not (rs.eof) then
		   rs("start")=3
           rs.update
		   rs.close
		   end if
		   
	num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="ȷ�Ϸ���"
		rs("content")=session("AdminName")&"����ȷ�Ϸ���"&id&"����"
		rs("price")=renum
		rs("jiegou")="ȷ�Ϸ����ɹ�"
    	rs.update
    	rs.close 
        response.Redirect("admin_mymission1.asp")

elseif action="hao" then

       Set rs=server.createobject("ADODB.RECORDSET")
       rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
       if not (rs.eof) then
           rs("start")=4
           rs.update
		   rs.close
	  end if
		   
		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"'",conn,3,3
           if not (rs.eof) then
		   rs("start")=4
           rs.update
		   rs.close
		   end if
		   
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
    	rs("num")=num
		rs("class")="�򷽺���"
		rs("content")=session("AdminName")&"�����򷽺���"&id&"����"
		rs("price")=renum
		rs("jiegou")="�򷽺����ɹ�"
    	rs.update
    	rs.close 
       response.Redirect("admin_mymission1.asp")

elseif action="haook" then

       Set rs=server.createobject("ADODB.RECORDSET")
      rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and start='4'",conn,3,3
       if not (rs.eof) then
	       fusername=rs("username")
		   price=rs("price")
		   jieducm_fb=rs("jieducm_fb")
           rs("start")=5
           rs.update
		   rs.close
		 else
		  Response.Write("<script>alert('��������!����״̬�ѷ����仯��');history.back();</script>")
          response.End()
		end if
				   
		   Set rs=server.createobject("ADODB.RECORDSET")
           rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"'",conn,3,3
           if not (rs.eof) then
		      jusername=rs("username")
		      rs("start")=5
              rs.update
			  rs.close
			end if

		   
fabu1=stupkou*int(jieducm_fb)
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select jifei,cunkuan,fabudian From "&jieducm&"_register where username='"&jusername&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+stupfajifei
rs("cunkuan")=rs("cunkuan")+price
rs("fabudian")=rs("fabudian")+fabu1
rs.update
rs.close

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select jifei From "&jieducm&"_register where username='"&fusername&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+stupfajifei
rs.update
rs.close
	
	
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=jusername
    	rs("num")=num
		rs("class")="�������"
		rs("content")=session("AdminName")&"���к�̨�����������"&id&"�������ѵõ�"&price&"Ԫ���"&fabu1&"��������"&stupfajifei&"����"
		rs("price")=price
		rs("jiegou")="������ɳɹ�"
    	rs.update
    	rs.close 
			   
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=fusername
    	rs("num")=num
		rs("class")="�������"
		rs("content")=session("AdminName")&"���к�̨�����������"&id&"���񣬽��ַ��ѵõ�"&price&"Ԫ���"&fabu1&"��������"&stupfajifei&"����"
		rs("price")=0
		rs("jiegou")="������ɳɹ�"
    	rs.update
    	rs.close 
       response.Redirect("admin_mymission1.asp")

end if

%>