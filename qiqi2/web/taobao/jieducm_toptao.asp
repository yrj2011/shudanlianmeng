<link  href="../css/header.css" rel="stylesheet" type="text/css"  />
<div align="center">
<div  style=" width:900px; margin-top:10px; BORDER-BOTTOM: #cccccc 1px dashed; BORDER-top: #cccccc 1px dashed;  BORDER-left: #cccccc 1px dashed;  BORDER-right: #cccccc 1px dashed; HEIGHT: 20px">
<div>
  <div align="left">亲爱的会员：<font color=#ff0000><%=session("useridname")%></font>
    <%if vip="1" then %>
    <img src="../images/jieducm_vip.gif"  alt="尊贵VIP,信誉值：<%call vipxinyu(session("useridname"))%>"  border="0"/>
    <% end if%>
    &nbsp;
    <%call tribename(tribeid)%>
    &nbsp;您好！您拥有存款：<font color=#ff0000>
<%
if (cunkuan)=0 then
response.Write("0.00")
elseif int(cunkuan)<1 then
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
else
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
end if
%>
    </font> 元 	&nbsp;
    发布点：<font color=#ff0000>
<%
if (fabudian)=0 then
response.Write("0.0")
elseif fabudian<1 then
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
else
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
end  if%>
    </font>点   &nbsp;
    积分：<font color=#ff0000><%=jifei%></font>点 
    <%if vip=1 then%>
VIP有效期：<FONT color=#ff0000><%=vipdate%></FONT>天；
VIP信誉值：<FONT color=#ff0000>
<%call vipxinyu(session("useridname"))%>
</FONT>
<%end if%>
<a href="../user/user.asp"> 站内信(<FONT color=#ff0000><%=weidu1%></FONT>)</a>
  </div>
</div>
        </div></div>

<TABLE width=910 align=center border=0>
  <TBODY>
  <TR>
    <TD>
      <DIV id=tjgg style="DISPLAY: none; WIDTH: 910px; LINE-HEIGHT: 100%; TEXT-ALIGN: left">
      <TABLE cellSpacing=0 cellPadding=0 width=910 border=0>
        <TBODY>
        <TR>
          <TD>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD width=108 height=62><IMG height=62  src="../images/left.gif" width=108 align=baseline useMap=#Map2 border=0></TD>
                <TD background=../images/t.gif>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD height=29></TD></TR>
                    <TR>
                      <TD height=22>
                        <DIV class=shell>
                        <DIV class=core id=div4>
						<%
Sql = "select Top 9 NewsPath,Title,ArticleID,ClassID,TitleFontColor,TitleFontType,UpdateTime from "&jieducm&"_Article where classid =2 order by Articleid desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	Response.Write("暂无信息")
Else
Do While Not Rs.Eof
 URL="../news.asp?/"&rs("Articleid")&".html"
						%>
                        <LI> <%response.Write("<a href='"+URL+"' target=""_blank"" title='"+rs("title")+"' >")%>
<%=left(rs("title"),15)%></A> 
                        </LI>
<%
Rs.MoveNext
Loop
End IF
rs.close
%>
                       </DIV></DIV></TD></TR>
                    <TR>
                      <TD height=11></TD></TR></TBODY></TABLE></TD>
                <TD width=28 height=62><IMG height=62 
                  src="../images/right.gif" width=28 useMap=#Map 
              border=0></TD>
              </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
			  
			  <SCRIPT>
function disclose()
{
(new Image).src="../s.asp";
  document.getElementById("tjgg").style.display="none";
  
}


//alert('444');
cl();

  function cl()
  {
  //disclose();
  //alert('DDDD');
  document.getElementById("tjgg").style.display="";
  }
 //alert('0');
</SCRIPT>
<SCRIPT>
function myGod(id,w,n){
	var box=document.getElementById(id),can=true,w=w||1500,fq=fq||10,n=n==-1?-1:1;
	box.innerHTML+=box.innerHTML;
	box.onmouseover=function(){can=false};
	box.onmouseout=function(){can=true};
	var max=parseInt(box.scrollHeight/2);
	new function (){
		var stop=box.scrollTop%18==0&&!can;
		if(!stop){
			var set=n>0?[max,0]:[0,max];
			box.scrollTop==set[0]?box.scrollTop=set[1]:box.scrollTop+=n;
		};
		setTimeout(arguments.callee,box.scrollTop%18?fq:w);
	};
};

myGod('div4',3000);

</SCRIPT>
<map name="Map" id="Map"><area shape="rect" onclick=disclose(); coords="4,4,29,30" href="#" /></map>