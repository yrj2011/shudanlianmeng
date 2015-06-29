<%

'==================================================
'过程名：ShowFriendLinks
'作  用：显示友情链接站点
'参  数：LinkType  ----链接方式，1为LOGO链接，2为文字链接
'       SiteNum   ----最多显示多少个站点
'       Cols      ----分几列显示
'       ShowType  ----显示方式。1为向上滚动，2为横向列表，3为下拉列表框
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
        strLink=strLink & "<div id=rolllink style=overflow:hidden;height:100;width:100><div id=rolllink1>"    '新增加的代码
	elseif ShowType=3 then
		strLink=strLink & "<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>友情文字链接站点</option>"
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
				strLink=strLink & "<td  align=left><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>"
				if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
				else
					strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
				end if
				strLink=strLink & "</a></td><td width=25>&nbsp;&nbsp;&nbsp;</td>"
			  else
				strLink=strLink & "<td><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>" & rsLink("SiteName")&"</a></td>"
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
        strLink=strLink & "</div><div id=rolllink2></div></div>"   '新增代码
	elseif ShowType=3 then
		strLink=strLink & "</select>"
	end if
	response.write strLink
	if ShowType=1 then call RollFriendLinks()    '新增代码
	rsLink.close
	set rsLink=nothing
end sub

'==================================================
'过程名：RollFriendLinks
'作  用：滚动显示友情链接站点
'参  数：无
'==================================================
sub RollFriendLinks()
%>
<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //克隆rolllink1为rolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //当滚动至rolllink1与rolllink2交界时
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink跳到最顶端
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //设置定时器
   rolllink.onmouseover=function() {clearInterval(MyMar)}//鼠标移上时清除定时器达到滚动停止的目的
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//鼠标移开时重设定时器
</script>
<%
end sub
%>


