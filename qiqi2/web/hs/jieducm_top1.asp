
<LINK href="../hs/main.css" type=text/css rel=stylesheet>
<%
dim restr
restr="index.asp"'�ļ�������ڵ�ǰ�ļ���ʱ
path= replace(replace(server.mappath(restr),server.mappath("/"),""),restr,"")
path1=replace(path,"\","")
call alipayto() 
Url=split(request.servervariables("script_name"),"/")
FileName=Url(ubound(Url))
File=Left(FileName,InstrRev(FileName,".")-1)
%>
<%
Url=split(request.servervariables("script_name"),"/")
FileName=Url(ubound(Url))
call alipayto() 
File=Left(FileName,InstrRev(FileName,".")-1)
%>
<script language="JavaScript" src="/hs/min.js"></script>
<script language="JavaScript" src="/hs/mask.js"></script>
<div id="sitenav"><div class="info"><a class="db_d"  onclick='$("#sitenav").slideToggle("slow")' style="cursor:pointer"></a><div class="db_z"><span style="font-weight:800;color: #ff0000">��ܰ��ʾ:</span> <MARQUEE onmouseover=this.stop()  onmouseout=this.start() scrollAmount=2 direction=left width="400" height=28><font color=#000000>���棺Ϊ��ÿ����Ա������Ͱ�ȫ��֤���������ֹС�������������½������һ�ο۳�5��������,����С�ž�ʹ����ҳ��½����ֹ������½С�ţ�������½����С�ţ���ֹ��������лл��������</font></a>
    </MARQUEE> </div><div class="db_y"><span style="color: #F6400">��ǰ����������<%=Application("OnLine")%>��&nbsp;&nbsp;��ӭ����</span>   <%if session("useridname")="" then%>
<A href="http://www.hushuawang.com/login.asp" target="_top">��½</A> <A  href="http://www.hushuawang.com/register.asp" target="_top"> ע��</A> 
<%else
Sql = "select fabudian,cunkuan,jifei,vip,tribeid from "&kuaishua&"_register where username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	response.Redirect("login.asp")				
Else
    fabudian=rs("fabudian")
	cunkuan=rs("cunkuan")
	jifei=rs("jifei")
	vip=Replace(Trim(rs("vip")),"'","''")
	tribeid=rs("tribeid")
end if
rs.close	
%>
 <A href="user/" target="_top"><SPAN><%=session("useridname")%></SPAN> </A>
 <%if vip="1" then%><img src="../images/vip.gif"  alt="���VIP" /><%end if%>   <A onClick="return confirm('ȷ���˳�������');"  href="http://www.hushuawang.com/user/exit.asp" target="_top">[�˳�]</A>  
 <% end if%>
 <A href="http://www.hushuawang.com/user/" target="_top"> �������� </A> <A href="#" title=��ӵ��ղؼ� onclick="window.external.addFavorite('<%=stupurl%>','<%=stupname%>')" >�ղر�վ</A> </div></div></div>
<div class="menu"><div class="logo"><A href="http://www.hushuawang.com" ><img src="http://www.hushuawang.com/uploadpic/201011921135324633.jpg"></A></div><div class="sj"><embed height="30" width="85" type="application/x-shockwave-flash" pluginspage="http://www.adobe.com/go/getflashplayer" quality="high" src="/hs/clock_4.swf?TimeZone=GMT0800" wmode="transparent" style="margin:0; padding:0;"></embed></div>
<div id="quick_menu">
      <ul>
        <li id="nav_index" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.hushuawang.com');return(false);">��Ϊ��ҳ</li>
        <li id="nav_favorite" onclick="window.open('http://shuqian.qq.com/post?from=3&title='+encodeURIComponent(document.title)+'&uri='+encodeURIComponent(document.location.href)+'&jumpback=2&noui=1','favit','');void(0)">QQ�ղؼ�</li>
        <li id="nav_qqgroup" onclick="window.location.href='http://www.hushuawang.com/'">ƽ̨QȺ</li>
      </ul><div> <a href=""id="showKF"><img src="http://www.hushuawang.com/hs/zx.gif"></a>
      <a href="javascript:alert('��ϲ��,���ѳɹ��л���������·!');window.location.href='http://www.hushuawang.com/'"><img src="/hs/06.png"/></a>  <img src="/hs/split.gif"/>  <a href="javascript:alert('��ϲ��,���ѳɹ��л�����ͨ��·!');window.location.href='http://www.hushuawang.com/'"><img src="/hs/07.png"/></a>  </div>
</div></div>

  <div class="nav">
  <div class="menul"></div><div class="menuc">
  
  <div class="menutab">
<ul>
            <li class="current" ><a href="/"><span>��վ��ҳ</span></a></li>
            <li><a href="/taobao/"><span>�Ա�����</span></a></li>
            <li><a href="/pai/"><span>���Ļ���</span></a></li>
            <li><a href="/you/"><span>�а�����</span></a></li>
            <li><a href="http://www.hushuawang.com//user/MoneyOrPush.asp"><span>���Ͷһ�</span></a></li>
            <li><a href="/user/login.asp"><span>����VIP</span></a></li>
            <li><a href="/union/"><span>�ƹ�׬Ǯ</span></a></li>
            <li><a href="/tribe/"><span>���佻��</span></a></li> 
            <li style="position:relative;"><a href="/new.asp"><span>��������</span></a><img src="/hs/Nddz_04.gif" style="position:absolute; top:-10px; right:0px; z-index:999;" /> </li>
           <li class="kuaisu" onclick="alert('������')"> </li>
          </ul>
        </div>
  
  
  </div><div class="menur"></div>
  </div>

 <div class="menud">
  <div class="menudl"></div><div class="menudc">
  <div class="userstau">�װ��ģ�<FONT color=#ff0000><%=session("useridname")%></FONT>&nbsp; <%if vip="1" then%><img src="../images/vip.gif"  alt="���VIP" /><%end if%>���ã�<font color="#FF0000"><%call tribename(tribeid)%> </font>����:
<FONT color=#ff0000><%=jifei%></FONT> <%call xinyu(jifei)%>  
������:<FONT color=#ff0000><%
if (fabudian)=0 then
response.Write("0.0")
else
fabudian=formatnumber((fabudian),1)
response.Write(fabudian)
end  if
%></FONT> ���:<FONT color=#ff0000><%
if (cunkuan)=0 then
response.Write("0.00")
else
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
end  if
%>
</FONT> Ԫ&nbsp;<font color="#FF0000"><%=weidu%> </font>��˽��վ����</div>
  <div class="menud2">����λ�ã�<a href="#" class="span2">���߻�ˢƽ̨</a> > <a href="http://www.778892.com/" class="span2">������ҳ</a></div>
  
  </div><div class="menudr"></div>
  </div>
	
	


