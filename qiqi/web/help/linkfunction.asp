<%

'==================================================
'��������ShowFriendLinks
'��  �ã���ʾ��������վ��
'��  ����LinkType  ----���ӷ�ʽ��1ΪLOGO���ӣ�2Ϊ��������
'       SiteNum   ----�����ʾ���ٸ�վ��
'       Cols      ----�ּ�����ʾ
'       ShowType  ----��ʾ��ʽ��1Ϊ���Ϲ�����2Ϊ�����б�3Ϊ�����б��
'==================================================
sub ShowFriendLinks(LinkType,SiteNum,Cols,ShowType)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if LinkType<>1 and LinkType<>2 then
		LinkType=1
	else
		LinkType=Cint(LinkType)
	end if
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	if Cols<=0 or Cols>20 then
		Cols=10
	end if
	if ShowType=1 then'
        strLink=strLink & "<div id=rolllink style=overflow:hidden;height:100;width:100><div id=rolllink1>"    '�����ӵĴ���
	elseif ShowType=3 then
		strLink=strLink & "<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>������������վ��</option>"
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "<table width='100%' cellSpacing='0'><tr align='center' >"
	end if
	
	sqlLink="select top " & SiteNum & " * from "&jieducm&"_FriendLinks where IsOK=True and LinkType=" & LinkType & " order by IsGood,id desc"
	set rsLink=server.createobject("adodb.recordset")
	rsLink.open sqlLink,conn,1,1
	if rsLink.bof and rsLink.eof then
		if ShowType=1 or ShowType=2 then
	  		for i=1 to SiteNum
				strLink=strLink & "<td>"			
				strLink=strLink & "</td>"
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' >"
				end if
			next
		end if
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			if ShowType=1 or ShowType=2 then
			  if LinkType=1 then
				strLink=strLink & "<td  align=left><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>"
				if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
				else
					strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
				end if
				strLink=strLink & "</a></td><td width=25>&nbsp;&nbsp;&nbsp;</td>"
			  else
				strLink=strLink & "<td><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>" & rsLink("SiteName")&"</a></td>"
			  end if
			  if i mod Cols=0 and i<SiteNum then
				strLink=strLink & "</tr><tr align='center' >"
			  end if
			else
				strLink=strLink & "<option value='" & rsLink("SiteUrl") & "'>" & rsLink("SiteName") & "</option>"
			end if
			rsLink.moveNext
		next
		if SiteCount<SiteNum and (ShowType=1 or ShowType=2) then
			for i=SiteCount+1 to SiteNum
				if LinkType=1 then
					strLink=strLink & "<td width='88'></td>"
				else
					strLink=strLink & "<td width='88'></td>"
				end if
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' >"
				end if
			next
		end if
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "</tr></table>"
	end if
	if ShowType=1 then
        strLink=strLink & "</div><div id=rolllink2></div></div>"   '��������
	elseif ShowType=3 then
		strLink=strLink & "</select>"
	end if
	response.write strLink
	if ShowType=1 then call RollFriendLinks()    '��������
	rsLink.close
	set rsLink=nothing
end sub

'==================================================
'��������RollFriendLinks
'��  �ã�������ʾ��������վ��
'��  ������
'==================================================
sub RollFriendLinks()
%>
<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //��¡rolllink1Ϊrolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //��������rolllink1��rolllink2����ʱ
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink�������
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //���ö�ʱ��
   rolllink.onmouseover=function() {clearInterval(MyMar)}//�������ʱ�����ʱ���ﵽ����ֹͣ��Ŀ��
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//����ƿ�ʱ���趨ʱ��
</script>
<%
end sub
%>


