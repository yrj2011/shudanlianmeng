
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="CLEAR: both; WIDTH: 930px; BACKGROUND-COLOR: #ffffff">
<DIV style="CLEAR: both; WIDTH: 100%">
<DIV style="CLEAR: both; WIDTH: 100%; HEIGHT: 45px">
<DIV class=c_01 style="WIDTH: 15%"><IMG src="../images/task_01.gif">������</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG src="../images/task_01.gif">������</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG src="../images/task_01.gif">����۸�</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG src="../images/task_01.gif">����ʱ��</DIV>
<DIV class=c_01 style="WIDTH: 20%"><IMG src="../images/task_01.gif">�ղؽ���</DIV>
<DIV class=c_01 style="CLEAR: right; WIDTH: 20%"><IMG src="../images/task_01.gif">��&nbsp;&nbsp;��</DIV></DIV></DIV></DIV></DIV>
<DIV id=missionset style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 950px"></DIV>
 <%	 
action=request("sort")
Sql = "select top 100  username,price,pid,bei,start,now,num,jieshou_num from "&jieducm&"_zhongxin where classid='6'  order by (num-jieshou_num) desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then				
 	if action="" then 
	  url="index.asp"
	else
	  url="index.asp?sort="&action&""
	end if
	 rs.pagesize=15
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
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
DO WHILE NOT rs.EOF AND RowCount>0
 bei=rs("bei")
 pid=rs("pid")
 now1=rs("now")
 username=rs("username")
 num=rs("num")
 jieshou_num=rs("jieshou_num")
 %>  
<DIV class=tt5>
<TABLE style="MARGIN: 3px" cellSpacing=2 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD width="15%"><%=pid%><BR><%=bei%></TD>
    <TD>
    <TD align=middle width="15%"><SPAN  style="Z-INDEX: 20;">
<%
Sql2 = "select jifei,qq,vip from "&jieducm&"_register where username='"&username&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then	
jifei=rs2("jifei")
vip=Replace(Trim(rs2("vip")),"'","''")	
qq=rs2("qq")
rs2.close
end if				
 %>
<%call isonline(username)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=username%>" title="����վ����Ϣ" target="_blank"><%=username%></a></SPAN>
<%if vip="1" then %><img src="../images/jieducm_vip.gif"  alt="���VIP,����ֵ��<%call vipxinyu(username)%>"  border="0"/><% end if%>
<br />
<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0" />
<BR>
<%call xinyu(jifei)%>
		 </TD>
    <TD align=middle width="15%"><img src="../images/jieducm_zf.gif" width="13" height="16"><font color="#FF0000"><strong>
	<%if num*stupcang<1 then
	response.Write("0")
	response.Write(num*stupcang)
	else
	response.Write(num*stupcang)
	end if
	%>
	</strong></font>��������</TD>
    <TD align=middle width="15%">
<%=now1%></TD>
    <TD style="PADDING-LEFT: 50px" width="20%">	
	<STRONG><%=jieshou_num%>/<%=num%></STRONG> 
	<DIV class=speedjieducm><DIV class=rate style="WIDTH: <%=jieshou_num/num*100%>%;  float: left;"></DIV>
	</DIV>
	</TD>
    <TD style="PADDING-LEFT: 50px" width="20%"> 
<%
if  username<>session("useridname") and jieshou_num<num then%>
<a title="���֣����������ɻ�ô��ͷ�����" style="CURSOR: pointer" onClick="showxiao('<%=pid%>','2')">
<img src=../images/online_admin.gif><span class=anniu>��������</span></a>
<%elseif jieshou_num=num then
response.Write("�����ѽ���")
else
response.Write("���ܽ��Լ�������")
end if%>
</TD></TD></TR></TBODY></TABLE></DIV>
<%
	  RowCount = RowCount - 1
	  i=i-1
      rs.MoveNext 
      Loop
%>  
<DIV class=tt5 >  <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></div>
