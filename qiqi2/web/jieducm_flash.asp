<%
set rs=server.createobject("adodb.recordset")
sql="select  * from "&jieducm&"_indexpic "
rs.open sql,conn,1,3
%>

 <SCRIPT language=javascript>
linkarr = new Array();
picarr = new Array();
textarr = new Array();

var swf_width=684;
var swf_height=214;
var files = "";
var links = "";
var texts = "";

linkarr[1] = "<%=rs("link1")%>";
picarr[1] = "<%=rs("pic1")%>";
textarr[1] = "<%=rs("title1")%>";
linkarr[2] = "<%=rs("link2")%>";
picarr[2] = "<%=rs("pic2")%>";
textarr[2] = "<%=rs("title2")%>";
linkarr[3] = "<%=rs("link3")%>";
picarr[3] = "<%=rs("pic3")%>";
textarr[3] = "<%=rs("title3")%>";
linkarr[4] = "<%=rs("link4")%>";
picarr[4] = "<%=rs("pic4")%>";
textarr[4] = "<%=rs("title4")%>";
for(i=1;i<picarr.length;i++){
  if(files=="") files = picarr[i];
  else files += "|"+picarr[i];
}

for(i=1;i<linkarr.length;i++){
  if(links=="") links = linkarr[i];
  else links += "|"+linkarr[i];
}

for(i=1;i<textarr.length;i++){
  if(texts=="") texts = textarr[i];
  else texts += "|"+textarr[i];
}

document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="images/bcastr3.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'">');
document.write('<embed src="images/bcastr3.swf" wmode="opaque" FlashVars="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'& menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</SCRIPT>
<%
rs.close
set rs=nothing
%>