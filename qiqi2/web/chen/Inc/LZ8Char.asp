<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name=keywords content="领智时代|领智网站系统">
<meta name="description" content="Design By LZ8">
<title><%=LZ8.WebSetting(2)%> - 控制面版</title>
<link rel="icon" href="Images/logo.ico" type="image/x-icon" />
<link rel="shortcut icon" href="Images/logo.ico" type="image/x-icon" />
<script src="images/admin.js"></script>
<script src="images/menu.js"></script>
<script language='JavaScript' type='text/JavaScript'>
function gocontent(){
	document.all("content").value=IframeID.document.body.innerHTML;
}
function AddItem(strFileName){
  var m = document.myform
  var arrName=strFileName.split('.');
  var FileExt=arrName[1];
  if (FileExt=='gif' || FileExt=='jpg' || FileExt=='jpeg' || FileExt=='jpe' || FileExt=='bmp' || FileExt=='png'){
	  m.IncludePic.checked=true;
	  m.IndexPicUrl.value=strFileName;
	  m.DefaultPicList.options[m.DefaultPicList.length]=new Option(strFileName,strFileName);
	  m.DefaultPicList.selectedIndex+=1;
  }
  if(m.UploadFiles.value==''){
	m.UploadFiles.value=strFileName;
  }else{
    m.UploadFiles.value=m.UploadFiles.value+"|"+strFileName;
  }
}
</script>
<link href="admin_css.css" rel="stylesheet" type="text/css">
</head>
<body>
<div align="center"><br>
<div class=menuskin id=popmenu onMouseOver="clearhidemenu();" onMouseOut="dynamichide(event)" style="Z-index:100"></div>
<style type=text/css>
.menuskin {
	BORDER: #666666 1px solid; VISIBILITY: hidden; FONT: 12px Verdana;
	POSITION: absolute; 
	BACKGROUND-COLOR:#EFEFEF;
	background-image:url("images/menubg.gif");
	background-repeat : repeat-y;
}
.menuskin A {
	PADDING-RIGHT: 10px; PADDING-LEFT: 32px; COLOR: black; TEXT-DECORATION: none;
}
#mouseoverstyle {
	BACKGROUND-COLOR: #C9D5E7; margin:2px; padding:0px; border:#597DB5 1px solid;
}
#mouseoverstyle A {
	COLOR: black;
}
.menuitems{
	margin:2px;padding:1px;word-break:keep-all; text-align : left;
}
</style>
<%
Session("AdminShowInfo") = Null
%>