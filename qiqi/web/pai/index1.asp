
<DIV style="CLEAR: both; WIDTH: 910px; BACKGROUND-COLOR: #f3f8fe"></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="CLEAR: both; WIDTH: 930px; BACKGROUND-COLOR: #ffffff">
<DIV style="CLEAR: both; WIDTH: 100%">
<DIV style="CLEAR: both; WIDTH: 100%; HEIGHT: 45px">
<DIV class=c_01 style="WIDTH: 15%"><IMG 
src="../images/task_01.gif">������</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG 
src="../images/task_01.gif">������</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG 
src="../images/task_01.gif">����۸�</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG 
src="../images/task_01.gif">���۷�ʽ</DIV>
<DIV class=c_01 style="WIDTH: 20%"><IMG 
src="../images/task_01.gif">����״̬</DIV>
<DIV class=c_01 style="CLEAR: right; WIDTH: 20%"><IMG 
src="../images/task_01.gif">��&nbsp;&nbsp;��</DIV></DIV></DIV></DIV></DIV>
<DIV >
<%	 
action=request.QueryString("sort")
  Sql = "select top 100  * from "&jieducm&"_zhongxin where classid='2'  order by start asc,id desc"
 Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then				
	if action="" then 
	  url="index.asp"
	else
	  url="index.asp?sort="&action&""
	end if
	 rs.pagesize=20
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
 start=rs("start")
 pid=rs("pid")
 now1=rs("now")
 price=rs("price")
 isprice=rs("isprice")
 username=rs("username")
 tips=rs("tips")
 jieducm_fb=rs("jieducm_fb")
 addfabu=0
if bei="�����ջ�����"  then
clas=1
else
clas=2
end if
 %>  
<DIV class=tt5 <%if clas=2 then%> style="BACKGROUND-COLOR: #cde0ee" <%end if%>>
<TABLE style="MARGIN: 3px" cellSpacing=2 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD width="15%"><%=pid%><BR><%=now1%></TD>
    <TD>
    <TD align=middle width="15%"><SPAN  style="Z-INDEX: 20; ">
	 <%
			  	Sql2 = "select * from "&jieducm&"_register where username='"&rs("username")&"'"
				Set Rs2 = Server.CreateObject("Adodb.RecordSet")
				Rs2.Open Sql2,conn,1,1
				if rs2.eof then
				else	
				jifei=rs2("jifei")
				qq=rs2("qq")
			  %>
<% call isonline(username)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=rs("username")%>" title="����վ����Ϣ" target="_blank"><%=rs("username")%></a>
<%end if%></SPAN><BR>
<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0" />
<br />
<%call xinyu(jifei)%>
		 </TD>
    <TD align=middle width="15%">
	ƽ̨���� <img src="../images/jieducm_zf.gif" width="13" height="16"><font color="#FF0000"><strong><%=Price%></strong></font>Ԫ<br><%=rs("isprice")%></TD>
    <TD align=middle width="15%">
<%=bei%><BR><%
if bei="�����ջ�����"  then
response.Write("<IMG alt=�����ջ����� src=../images/jieducm_xuni.gif>")
else
response.Write("<IMG alt=�ӳ��ջ� src=../images/jieducm_shiwu.gif>")
end if%>������<font color="#FF0000"><strong><%
call jieducm_fadian(jieducm_fb,addfabu)
 %></strong></font>��</TD>
    <TD style="PADDING-LEFT: 50px" width="20%"><B   <%if start=0 then%> style="COLOR: red" <%else %>style="COLOR: #006600" <%end if%>>	
	<%call zhuangtai(start)%>
      </B><br /> <%if tips<>"" then
	response.Write("<br>(<a  title='"&tips&"'>")
	response.Write(left(tips,6))
	response.Write("...)")
	end if%></TD>
    <TD style="PADDING-LEFT: 50px" width="20%"> 
	<%
	if start="0" and rs("username")<>session("useridname")then
response.Write("<img src=../images/online_admin.gif>")%>
<a title="���֣����������ɻ�ô��ͷ�����" style="CURSOR: pointer" onClick="showxiao('<%=rs("pid")%>','2')">
<span class=anniu>��������</span></a>
<%elseif (start=1 or start=2 or start=3 or start=4) then
response.Write("�ѽ���")
elseif start=5 then
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
  </div>
