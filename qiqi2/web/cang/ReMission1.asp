 
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 5px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 125px; COLOR: #006600; TEXT-ALIGN: center">����ID</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">������</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">����۸�</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">��Ʒ��Ϣ</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 140px; COLOR: #006600; TEXT-ALIGN: center">����ʱ�� 
</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 120px; COLOR: #006600; TEXT-ALIGN: center">״&nbsp;&nbsp;̬</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 160px; COLOR: #006600; TEXT-ALIGN: center">��&nbsp;&nbsp;��</DIV></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid;BACKGROUND-COLOR: #ffffff">
<%
action=request.QueryString("sort")
if action="" then
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"'  order by start asc ,id desc"
elseif action="6" then
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"'  order by start asc"
else
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"' and start='"&action&"' order by start  asc"
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if rs.eof then
 response.write("���޽�����Ϣ")
 end if
if not rs.eof then				
 	if action="" then
	 url="remission.asp"
	else
	 url="remission.asp?sort="&action&""
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
if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
DO WHILE NOT rs.EOF AND RowCount>0
pid=rs("pid")
now1=rs("now")
prourl=rs("prourl")
Shopkeeper=rs("Shopkeeper")
num=rs("num")
start=rs("start")
id=rs("id")
Sql2 = "select * from "&jieducm&"_zhongxin where pid='"&pid&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then
play=rs2("play")
nowf=rs2("now")
usernamef=rs2("username")
tips=rs2("tips")
bei=rs2("bei")
end if
 %>	  
<DIV class=missionbg style="WIDTH: 98%; PADDING-TOP: 10px; HEIGHT: 70px">
<DIV style="FLOAT: left; WIDTH: 125px; LINE-HEIGHT: 150%; TEXT-ALIGN: left"><%=pid%>
<BR><%=bei%>
<BR>
<%=nowf%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; TEXT-ALIGN: center">
<DIV style="OVERFLOW: hidden; WIDTH: 100%;  TEXT-ALIGN: center">
<%	Sql = "select jifei,qq,vip from "&jieducm&"_register where username='"&usernamef&"'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql,conn,1,1
if not rs1.eof then
jifei=rs1("jifei")
qq=rs1("qq")
vipf=Replace(Trim(rs1("vip")),"'","''")
rs1.close
end if
 %>
<%call isonline(usernamef)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=usernamef%>" title="����վ����Ϣ" target="_blank"><%=usernamef%></a> 
<%if vipf="1" then %><img src="../images/jieducm_vip.gif"  alt="���VIP,����ֵ��<%call vipxinyu(usernamef)%>"  border="0"/><% end if%>
<br>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes">
<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0"  />
</a>
<%call xinyu(jifei)%><br>
</DIV>

</DIV>
<DIV style="FLOAT: left; WIDTH: 115px;  TEXT-ALIGN: center"><font color=#ff6600>ƽ̨����<br>
  <img src="../images/jieducm_zf.gif" width="13" height="16"><%=stupcang%>��������</font><br>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; LINE-HEIGHT: 100%; TEXT-ALIGN: center">
    
  <label>
  <input name="<%=pid%>" type="hidden" id="<%=pid%>" size="9" value="<%=prourl%>">
  <INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=������Ʒ��ַ>
  <br>
  </label>
  <a onClick="javascript:clipboardData.setData('text','<%=tips%>');alert('��ǩ�ѳɹ�����');return false;" href="#">�����Ʊ�ǩ��</a> </DIV>
<DIV style="FLOAT: left; WIDTH: 140px; LINE-HEIGHT: 150%; TEXT-ALIGN: center"><%=now1%><BR>
<font color="#FF0000">�����Ա��ţ�</font><br><%=num%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 120px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=1 then
response.Write("<B style='COLOR: red'>�ȴ��ղ�...  </B>")
elseif start=2 then
response.Write("<B style='COLOR:#006600'>������ղ� </B>")
end if
%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 160px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=1 then
sss= DateDiff( "n",now1, Now())
if 20-sss>=1 then
%>
<a href="#1" onClick="javascript:showDialog('��ȷ�����Ѿ��ղ�����',1,7,'jieducm_ok.asp?id=<%=pid%>')" >
<SPAN class=anniu>�Ѿ��ղ�</SPAN>
</A>
<BR> 
ʣ��<span style="color: #ff0000"><span id="min1"><%=20-sss%></span> </span>���ӿ��ղ�
<%else

sqlr="update "&jieducm&"_zhongxin set  jieshou_num=jieshou_num-1 where  pid='"&pid&"'  and classid='6'"
conn.execute sqlr

sqlrs="delete  from "&jieducm&"_jieshou where id='"&id&"'  and classid='6'"
conn.execute sqlrs

end if%>

<BR>��ϵ�������ɼ�ʱ��<BR>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ�˳�����������',1,7,'jieducm_remis.asp?id=<%=id%>')" title="��ȷ��Ҫ�˳�����������">
<SPAN class=anniu>�˳�����</SPAN></A>
<%end if%>

</DIV>
</DIV>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 3px; BORDER-BOTTOM: #06314a 1px dashed;">
</DIV>   
<%
RowCount = RowCount - 1
i=i-1
rs.MoveNext 
Loop
%>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center">
 <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></DIV></DIV>
 