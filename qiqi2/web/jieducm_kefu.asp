<SCRIPT language=javascript type=text/javascript>
function Button2_onclick() {
fsdx.style.display='';
}
function Button3_onclick() {
kfqq.style.display='';
}
</SCRIPT>

<style type="text/css">
.qqStyle{width:100%; clear:both; margin:5px 0 0 5px; height:30px; line-height:30px; overflow:hidden;font-size:20px; font-weight:bolder;}
.qq{float:left; width:105px; height:52px; margin:5px 0 0 5px; overflow:hidden;}
.qq1{float:left; width:155px; height:40px; margin:5px 0 0 5px; overflow:hidden;}
</style>
<DIV id=kfqq style="DISPLAY: none">
<DIV id=kfqq_top>
<DIV id=kfqq_top_left><%=stupname%>�ͷ�����</DIV>
<DIV id=kfqq_top_right></DIV></DIV>
<DIV id=kfqq_bottom>
<DIV id=kfqq_bottom_top>
<DIV id=kfqq_bottom_top_top>
<DIV id=kfqq_bottom_top_top_top>�ͷ�������ʱ�� 9��00-18��00 ����12��00-14��00��Ϣ�Է�ʱ��</DIV>
<DIV id=kfqq_bottom_top_top_bottom>
<table cellspacing="1" width="100%" align="center" cellpadding="2" border="0">
		<tr>
			<td valign="top">
			<div class="qqStyle">ҵ��ָ����</div>
	<div style="width:100%; clear:both; overflow:hidden;border-bottom:#DFDFDF 1px dashed;">
			    <%
			  	Sql = "select Top 10 * from "&jieducm&"_qq where class=1 order by ID desc"
				Set rs = Server.CreateObject("Adodb.RecordSet")
				rs.Open Sql,conn,1,1
				IF rs.Eof Then
					
				Else
			
					Do While Not rs.Eof
						
			  %>
                       <div class="qq">
		   
		 <a target="_blank"  href="http://wpa.qq.com/msgrd?v=3&uin=<%=rs("qq")%>&site=qq&menu=yes" ><img src="http://wpa.qq.com/pa?p=1:<%=rs("qq")%>:41"  border="0"  /> <%=rs("qq")%><BR><%=rs("num")%></A><br />
                       </div>
                    <%
			  	rs.MoveNext
				Loop
				End IF
				rs.close
			  %>
              
                    
			</div>

			<div class="qqStyle">����רԱ��</div>
			<div style="width:100%; clear:both; overflow:hidden;border-bottom:#DFDFDF 1px dashed;">
			    <%
			  	Sql = "select Top 10 * from "&jieducm&"_qq where class=2 order by ID desc"
				Set rs = Server.CreateObject("Adodb.RecordSet")
				rs.Open Sql,conn,1,1
				IF rs.Eof Then
					
				Else
			
					Do While Not rs.Eof
						
			  %>
                       <div class="qq">
           <a target="_blank"  href="http://wpa.qq.com/msgrd?v=3&uin=<%=rs("qq")%>&site=qq&menu=yes" ><img src="http://wpa.qq.com/pa?p=1:<%=rs("qq")%>:41"  border="0"  /> <%=rs("qq")%><BR><%=rs("num")%></A><br />
                       </div>
 <%
			  	rs.MoveNext
				Loop
				End IF
				rs.close
			  %>
			</div>
			
			<div class="qqStyle">Ͷ�߽��飺</div>
			<div style="width:100%; clear:both; overflow:hidden;border-bottom:#DFDFDF 1px dashed;">
			   <div class="qq">
           <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1662424999&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:1662424999:41"  border="0"  /><BR /> 1662424999<BR />����רԱ</A><br />
                       </div>
					   <div class="qq">
           <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1662424999&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:1662424999:41"  border="0"  /><BR />1662424999<BR />Ͷ�߽���</A><br />
                       </div> </div>
					   <div style="width:100%; clear:both; overflow:hidden;border-bottom:#DFDFDF 1px dashed;">
					   <div class="qqStyle">�ٷ�QQȺ��372860439</div>
					   <div class="qqStyle">�������ߣ�13892299051</div>
			</div></div>
            </td>
		</tr>
	</table></DIV></DIV></DIV><DIV id=kfqq_bottom_bottom>������Ѷ��ʱ�Ự����ʧ�ܣ����Ϊ���ѣ�</DIV>
	 <div align="center"><A onClick="kfqq.style.display='none';" href="javascript:void(0)"><img src="http://www.renrenlai.com/images/gbi.gif"  border="0"  /></A>
					   </div>
	</DIV></DIV>
