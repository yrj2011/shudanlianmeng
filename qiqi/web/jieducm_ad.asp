
<%
sub gglei(wei,lei)

set rsgg=conn.execute("select * from "&jieducm&"_Advertising where HL_ggzt=0 and  HL_ggwei="&cint(wei)&" order by HL_ggID")
if rsgg.bof and rsgg.eof then
else
do while not rsgg.eof
   set rslei=conn.execute("select * from "&jieducm&"_ggwei where HL_id="&cint(rsgg("HL_gglei")))
   height=rslei("HL_height")
   width=rslei("HL_width")
   rslei.close
   set rslei=nothing
   if lei="首页广告位二" then
    listheight=4
   else
    listheight=3
   end if
%>
<div style="float:left;height:<%=height%>px;width:<%=width%>px;text-align:center;padding:<%=listheight%>px 1px <%=listheight%>px 1px;">
<%
if rsgg("HL_ggtu")="True" then
%><a href="<%=rsgg("HL_home")%>" target="_blank"><img src="../<%=rsgg("HL_ggImg")%>" border="0" width="<%=width%>px" height="<%=height%>px" ></a><%else%><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="<%=width%>" height="<%=height%>">
          <param name="movie" value="<%=rsgg("HL_ggImg")%>">
          <param name="quality" value="high">
          <embed src="<%=rsgg("HL_ggImg")%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="<%=width%>" height="<%=height%>"></embed>
        </object><%
end if
%></div><%
rsgg.movenext
loop
end if

end sub

sub ggwei(wei)
set rsgw=conn.execute("select * from "&jieducm&"_ggwei where HL_title='"&wei&"'")
if rsgw.bof and rsgw.eof then
else
call gglei(rsgw("HL_id"),wei)
'response.Write(rs("hl_Id"))
end if
'response.Write(wei)
end sub
%>

