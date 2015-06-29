<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
if session("AdminName")="" then
   response.redirect("admin_login.asp")
end if%>
<style type="text/css">
<!--
body {
	background-color: #CED7F7;
}
body,td,th {
	font-size: 12px;
}
.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
-->
</style>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>

<link href="inc/url.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE3 {color: #FF3300}
-->
</style>
<head>
<title>您的信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body text="#000000">
<%if request("action")="tj" or request("action")="xg" then

if request("action")="xg" then


set rsxg=conn.execute("select * from "&jieducm&"_Advertising where id="&cint(request("id")))

end if

%>
<form action="?action=HL_gg<%=request("action")%>&id=<%=request("id")%>" method="post" name="cmsfrm">
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr>
    <td colspan="3" bgcolor="#A6B5F0"><table cellpadding="2" cellspacing="1" border="0" width="100%" align="center" class="a2">
      <tr>
        <td height="25" colspan="3" align="center" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><b>系统栏目添加</b></td>
      </tr>
      <tr class="a4">
        <td height="23" align="right" bgcolor="#CED7F7">广告标题：</td>
        <td bgcolor="#CED7F7"><input name="HL_Title" type="text" id="HL_Title" value="<%
		if request("action")="xg" then
		response.write rsxg("HL_Title") 
		end if
		%>"></td>
        <td bgcolor="#CED7F7"><input name="HL_ggtu" type="radio" value="True" <%
		if request("action")="xg" then
		   if rsxg("HL_ggtu")="True" then
		     response.write "checked"
		   end if
		 else
		     response.write "checked"
		end if
		%>>
图片广告
  <input name="HL_ggtu" type="radio" value="False" <%
		if request("action")="xg" then
		  if rsxg("HL_ggtu")="False" then
		   response.write "checked"
		  end if 
		end if
		%>>
flash广告</td>
      </tr>
      <tr class="a4">
        <td width="15%" height="28" align="right" bgcolor="#CED7F7">所属广告位：</td>
        <td width="20%" bgcolor="#CED7F7"><div><select name="ggwei1" id="ggwei1" size="1" style="position:absolute; left: 198px; top: 70px; width:120px; height:20px;clip: rect(0 120 21 100)" onChange="ggwei.value=ggwei1.value;ggwei.select()">
		<%
		
		set rsggwei=conn.execute("Select * from "&jieducm&"_ggwei where HL_Lei=0")
		if rsggwei.bof and rsggwei.eof then
		else
		do while not rsggwei.eof
		response.write "<option value="""&rsggwei("HL_title")&""" >"&rsggwei("HL_title")&"</option>"
		rsggwei.movenext
		loop
		end if
		rsggwei.close
		set rsggwei=nothing
		%>
<!--                                <option   value="首页广告位一">首页广告位一</option>
<option   value="首页广告位二">首页广告位二</option>-->
                              </select><input type="text" style="position:absolute; left:199px; top:71px;width:102px; height:21px;" name="ggwei" value="<%
		if request("action")="xg" then
		call ggweitile(rsxg("HL_ggwei"),"HL_title")
		else
		response.write "请选择或输入！" 
		end if
		%>" id="ggwei" onFocus="if(value=='请选择或输入！') {value=''}" onBlur="if 
(value=='') {value='请选择或输入！'}">
        </div></td>
        <td width="65%" bgcolor="#CED7F7">          如果没有此广告位,请手动输入.</td>
      </tr>
      <tr class="a4">
        <td height="28" align="right" bgcolor="#CED7F7">广告类别：</td>
        <td bgcolor="#CED7F7"><div>
          <input type="text" style="position:absolute; left:198px; top:102px;width:102px; height:21px;" name="ggflei" value="<%
		if request("action")="xg" then
		call ggweitile(rsxg("HL_gglei"),"HL_title")
		else
		response.write "请选择或输入！" 
		end if
		%>" id="ggflei" onFocus="if(value=='请选择或输入！') {value=''}" onBlur="if 
(value=='') {value='请选择或输入！'}"><select name="ggflei1" id="ggflei1" size="1" style="position:absolute; left: 195px; top: 102px; width:120px; height:20px;clip: rect(0 120 21 100)" onChange="ggflei.value=ggflei1.value;ggflei.select()">
		<%
		
		set rsggwei=conn.execute("Select * from "&jieducm&"_ggwei where HL_Lei=1")
		if rsggwei.bof and rsggwei.eof then
		else
		do while not rsggwei.eof
		response.write "<option value="""&rsggwei("HL_title")&""" >"&rsggwei("HL_title")&"</option>"
		rsggwei.movenext
		loop
		end if
		rsggwei.close
		set rsggwei=nothing
		%>
                              </select>
        </div></td>
        <td bgcolor="#CED7F7">如需手动输入,名称与参数用&quot;<span class="STYLE1">|</span>&quot;号隔开,宽度与高度用&quot;<span class="STYLE1">*</span>&quot;号隔开.例:<span class="STYLE3">小栏|90*80</span></td>
      </tr>
      <tr class="a4">
        <td height="23" align="right" bgcolor="#CED7F7">广告图片：</td>
        <td colspan="2" bgcolor="#CED7F7"><input name="Upfile" type="text" value="<%
		if request("action")="xg" then
		response.write rsxg("HL_ggImg") 
		end if
		%>">
        <input type="button" name="Submit" value="上传图片" onClick="window.open('../upload_flash.asp?formname=cmsfrm&editname=Upfile&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
        
            &nbsp;</td>
      </tr>
      <tr class="a4">
        <td height="23" align="right" bgcolor="#CED7F7">链接地址：</td>
        <td colspan="2" bgcolor="#CED7F7"><input name="HL_Urltj" type="text" id="HL_Urltj" value="<%
		if request("action")="xg" then
		HL_home=Split(rsxg("HL_home"),"//")
		response.write rsxg("HL_home")
		end if
		%>" size="46">
          (必须以http://开头)</td>
      </tr>

      <tr class="a4">
        <td height="23" align="right" bgcolor="#CED7F7">&nbsp;</td>
        <td colspan="2" bgcolor="#CED7F7"><input type="submit" name="button" id="button" value="添加广告"></td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
<%
end if
if request("action")="" then
call adminlist()
response.write "</br>"
%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" >
    <tr>
      <td colspan="3" bgcolor="#A6B5F0"><table cellpadding="2" cellspacing="1" border="0" width="100%" align="center" class="a2">
          <tr>
            <td height="25" colspan="6" align="center" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><b>系统栏目信息</b></td>
          </tr>
          <tr class="a4">
            <td width="6%" height="23" align="center" bgcolor="#CED7F7">选择</td>
            <td width="28%" align="center" bgcolor="#CED7F7">广告名称</td>
            <td width="16%" align="center" bgcolor="#CED7F7">图片</td>
            <td width="14%" align="center" bgcolor="#CED7F7">广告位</td>
            <td width="12%" align="center" bgcolor="#CED7F7">广告类别</td>
            <td width="24%" align="center" bgcolor="#CED7F7">管理选项</td>
        </tr>
<form action="HL_admin_gg.asp?action=del" method="post" name="delform">
<INPUT NAME="mode" TYPE="hidden">
<INPUT NAME="page" TYPE="hidden" value="<%=request("page")%>">
<%

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from "&jieducm&"_Advertising order by id",conn,1,3

dim MaxPerPage,disturl
disturl="HL_admin_gg.asp"
MaxPerPage=8
text="0123456789"
rs.PageSize=MaxPerPage
for i=1 to len(request("page"))
   checkpage=instr(1,text,mid(request("page"),i,1))
   if checkpage=0 then
      exit for 
   end if
next
If checkpage<>0 then
      If NOT IsEmpty(request("page")) Then
        CurrentPage=Cint(request("page"))
        If CurrentPage < 1 Then CurrentPage = 1
        If CurrentPage > rs.PageCount Then CurrentPage =rs.PageCount
      Else
        CurrentPage= 1
      End If
      If not rs.eof Then rs.AbsolutePage = CurrentPage end if
Else
   CurrentPage=1
End if

if rs.bof and rs.eof then
else
i=0
do while not rs.eof 
%>               
            <tr>
            <td width="6%" height="23" align="center" bgcolor="#CED7F7"><input type="checkbox" name="id" id="id" value="<%=rs("id")%>"></td>
            <td width="28%" bgcolor="#CED7F7"><a href="<%=rs("HL_home")%>" target="_blank"><%=rs("HL_title")%></a></td>
            <td width="16%" bgcolor="#CED7F7"><%
			if rs("HL_ggtu")="False" then
			Response.Write "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width=""140"" height=""39"">"
Response.Write "          <param name=""movie"" value=../"& rs("HL_ggimg") &">"
Response.Write "          <param name=""quality"" value=""high"">"
Response.Write "          <embed src==../"&rs("HL_ggimg")&" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""140"" height=""39""></embed>"
Response.Write "</object>"
else
response.write "<a href=../"&rs("HL_ggImg")&" target=_blank><img src=../"&rs("HL_ggImg")&" width=""139"" height=""39"" border=""0""></a>"
end if
%></td>
            <td align="center" bgcolor="#CED7F7"><%call ggweitile(rs("HL_ggwei"),"HL_title")%></td>
            <td align="center" bgcolor="#CED7F7"><%call ggweitile(rs("HL_gglei"),"HL_title")%></td>
            <td width="24%" align="center" bgcolor="#CED7F7"><%
		if rs("HL_ggzt")=False then
		%><a href="?action=del&mode=tingyong&id=<%=rs("id")%>&User_zt=1">停用</a><%
		else
		%>
		<a href="?action=del&mode=tingyong&id=<%=rs("id")%>&User_zt=0"><font color="red">启用</font></a>
		<%
		end if
		%> | <a href="?action=xg&id=<%=rs("id")%>">修改</a> | <a href="?action=del&mode=del&id=<%=rs("id")%>&page=<%=request("page")%>" onClick="{if(confirm('此操作会删除当前分类及下层所有分类。确定这样做吗？')){return true;}return false;}">删除</a> </td>
        </tr>
 <%
 i=i+1 
 if i>=MaxPerPage then exit do end if
     rs.movenext
	loop
  end if
 %> <script language="javascript">

function selectIt(action){

var testform=document.getElementById("delform");

for(var i=0 ;i<testform.elements.length;i++){

if(testform.elements[i].type=="checkbox"){

e=testform.elements[i];

e.checked=(action=="selectAll")?1:(!e.checked);

}

}

}
function dellist() {

		delform.mode.value = "del";
		delform.submit();

}
function tingyong(){
		delform.mode.value = "tingyong";
		delform.submit();
}
function qiyong(){
		delform.mode.value = "qiyong";
		delform.submit();
}
</script>

             <tr>
              <td height="23" align="center" bgcolor="#CED7F7">&nbsp;</td>
              <td colspan="5" bgcolor="#CED7F7"><a href="javascript:;" onClick="return selectIt()">全选/反选</a>  <a href="javascript:tingyong()" onClick="{if(confirm('此操作会停用所有已选择的信息。确定这样做吗？')){return true;}return false;}">停用所选</a> <a href="javascript:qiyong()" onClick="{if(confirm('此操作会启用所有已选择的信息。确定这样做吗？')){return true;}return false;}">启用所选</a> <a href="javascript:dellist()" onClick="{if(confirm('此操作会删除所择选的所有信息。确定这样做吗？')){return true;}return false;}">删除所选信息</a></td>
            </tr>
			</form>
             <tr>
               <td height="23" colspan="6" align="center" bgcolor="#CED7F7"><%
			      If currentpage > 1 Then
      response.write "<a href='"&disturl&"?&page="+cstr(1)+"'>首页</a><font color='#ffffff'><b>-</b></font>"  
      Response.write "<a href='"&disturl&"?&page="+Cstr(currentpage-1)+"'>上一页</a><font color='#ffffff'><b>-</b></font>"
   Else
      Response.write "首页<font color='#000000'><b>-</b></font>"
      Response.write "上一页<font color='#000000'><b>-</b></font>"      
   End if
   
   If currentpage < rs.PageCount Then
      Response.write "<a href='"&disturl&"?page="+Cstr(currentPage+1)+"'>下一页</a><font color='#000000'><b>-</b></font>"
      Response.write "<a href='"&disturl&"?page="+Cstr(rs.PageCount)+"'>尾页</a>&nbsp;&nbsp;"
   Else
      Response.write "下一页&nbsp;"
      Response.write "尾页&nbsp;"       
   End if
   Response.write "共" & "<font color=red>" & Cstr(rs.RecordCount) & "</font>" & "<font color='#000000'>条留言</font>&nbsp;&nbsp;"
   Response.write "当前页:" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font>&nbsp;&nbsp;" 
   
 response.write "<select name=""select"" onchange=""window.location=this.value;"">"
 if request("page")="" then
 page=1
 else
 page=request("page")
 end if
 for j= 1 to Cstr(rs.pagecount)
 response.write "<option value=?Page="&j
 if cint(page)=cint(j) then
 response.write " selected"
 end if
 response.write ">第"& j &"页</option>"
 Next
 response.write " </select>"
 %></td>
             </tr> 
			        
      </table></td>
    </tr>
</table>
<%

end if
if request("action")="ORder" then
call adminlist()
response.write "</br>"
%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" >
    <tr>
      <td colspan="3" bgcolor="#A6B5F0"><table cellpadding="2" cellspacing="1" border="0" width="100%" align="center" class="a2">
          <tr>
            <td height="25" colspan="5" align="center" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><b>排序管理</b></td>
          </tr>

<%


set rs = conn.execute("select Max(HL_ggID) From "&jieducm&"_Advertising where HL_ggwei="&cint(request("id")))
         Hl_Home=rs(0)
    if isnull(Hl_Home) then
         Hl_Home=1
    end if
    rs.close
   set rs=nothing    


set rs=conn.execute("select * from "&jieducm&"_Advertising where HL_ggwei="&cint(request("id"))&" order by HL_ggid")
if rs.bof and rs.eof then
else
j=0
do while not rs.eof 
j=j+1
%>
            
            <tr>
			<td width="20%" bgcolor="#CED7F7">&nbsp;&nbsp;<%=rs("HL_title")%></td>
            <td width="20%" height="23" align="center" bgcolor="#CED7F7"><%
			if rs("HL_ggtu")="False" then
			Response.Write "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width=""140"" height=""39"">"
Response.Write "          <param name=""movie"" value=../'"& rs("HL_ggimg") &"'>"
Response.Write "          <param name=""quality"" value=""high"">"
Response.Write "          <embed src=../'"&rs("HL_ggimg")&"' quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""140"" height=""39""></embed>"
Response.Write "</object>"
else
response.write "<a href=../"&rs("HL_ggImg")&" target=_blank><img src=../"&rs("HL_ggImg")&" width=""139"" height=""39"" border=""0""></a>"
end if
%></td>
			<td bgcolor="#CED7F7" width="24%" align="center"><%call ggweitile(rs("HL_gglei"),"HL_title")%></td>
            <%
			if rs("HL_ggid")>1 then
			
			set trs=conn.execute("select count(HL_ggid) From "&jieducm&"_Advertising where HL_ggwei="&cint(request("id"))&" and HL_ggid<"&rs("HL_ggid")&"") 
		UpMoveNum=trs(0) 
		if isnull(UpMoveNum) then UpMoveNum=0 
		trs.close
		if UpMoveNum>0 then
         response.write "<form action=""?action=ORderp"" method=""post"" name=""cmsfrm"">"
	    	response.write "<td width=""18%"" align=""center"" bgcolor=""#CED7F7"">"  
             response.write "<select name=""HL_ggid"">"
             response.write "<option value=0>向上移动</option>"
				for i=1 to UpMoveNum
				response.write "<option value="&i&" >"&i&"</option>"
				
				Next
              response.write "</select>&nbsp;&nbsp;"
			  response.write "<input type=hidden name=ggid value="&rs("id")&">"
			  response.write "<input type=submit name=Submit value=修改></td>"
			  response.write "</form>"	
			end if			  		
		
		else
		response.write "<td width=""18%"" align=""center"" bgcolor=""#CED7F7""></td>"
		end if
		
			  

			if rs("HL_ggid")<Hl_Home then
			
			set trs=conn.execute("select count(HL_ggid) From "&jieducm&"_Advertising where HL_ggwei="&cint(request("id"))&" and HL_ggid>"&rs("HL_ggid")&"") 
		UpMoveNum=trs(0) 
		if isnull(UpMoveNum) then UpMoveNum=0 
		trs.close
		if UpMoveNum>0 then
		response.write "<form action=""?action=ORderb"" method=""post"" name=""cmsfrm"">" 
		response.write "<td width=""20%"" align=""center"" bgcolor=""#CED7F7"">"
             response.write "<select name=""HL_ggid"">"
             response.write "<option value=0>向下移动</option>"
				for i=1 to UpMoveNum
				response.write "<option value="&i&" >"&i&"</option>"
				Next
              response.write "</select>&nbsp;&nbsp;"
			  response.write "<input type=hidden name=ggid value="&rs("id")&">"
			  response.write "<input type=submit name=Submit value=修改></td>"
			  response.write "</form>"	
			end if	
		
		
		else
		response.write "<td width=""25%"" align=""center"" bgcolor=""#CED7F7""></td>"
		end if
		
			%>

          </tr>
	
 <%
     rs.movenext
	loop
  end if
 %>         
      </table></td>
    </tr>
</table>
<%
end if
sub adminlist()
%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" >
    <tr>
      <td colspan="3" bgcolor="#A6B5F0"><table cellpadding="2" cellspacing="1" border="0" width="100%" align="center" class="a2">
          <tr>
            <td width="8%" height="25" align="center" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><strong><a href="HL_admin_gg.asp">广告管理</a></strong></td>
            <td width="92%" background="images/titlebg.gif" bgcolor="#CED7F7" class="a1"><%
			
			set rsorder=conn.execute("select * from "&jieducm&"_ggwei where HL_lei=0")
			if rsorder.bof and rsorder.eof then
			else
			  do while not rsorder.eof 
			  
			  response.write "<a href=HL_admin_gg.asp?action=ORder&id="&rsorder("HL_id")&" >"
			  response.write rsorder("HL_Title")&"排序"
			  response.write "</a>"
			  response.write "&nbsp;&nbsp;"
			  
			  rsorder.movenext
			  loop
			end if
			%></td>
          </tr>
          
      </table></td>
    </tr>
</table>
<%
end sub
%>
</body>
</html>
<%
if request("action")="HL_ggtj" or request("action")="HL_ggxg" then

'**********************
'添加广告位及广告分类
'**********************
ggflei=request("ggflei")
ggwei=request("ggwei")





if ggflei="请选择或输入！" or ggwei="请选择或输入！" then
response.Write "<script LANGUAGE='javascript'>alert('您没有选择广告位或广告类型！');history.go(-1);</script>"
response.End()
end if


if ggwei<>"请选择或输入！" then

set rs=conn.execute("Select HL_title from "&jieducm&"_ggwei where HL_title='"&ggwei&"'")
if rs.bof and rs.eof then
   set rstj=server.CreateObject("adodb.recordset")
    rstj.open "select * from "&jieducm&"_ggwei",conn,1,3
	rstj.addnew
	rstj("HL_title")=ggwei
	rstj("HL_Lei")=0
	rstj.update
	rstj.close
	set rstj=nothing
end if
rs.close
set rs=nothing

end if


if ggflei<>"请选择或输入！" then
set rs=conn.execute("Select HL_title from "&jieducm&"_ggwei where HL_title='"&ggflei&"'")
if rs.bof and rs.eof then
gglei=Split(ggflei,"|")
ggcc=Split(gglei(1),"*")
   set rstj=server.CreateObject("adodb.recordset")
    rstj.open "select * from "&jieducm&"_ggwei",conn,1,3
	rstj.addnew
	rstj("HL_Title")=ggflei
	rstj("HL_width")=cint(ggcc(0))
	rstj("HL_height")=cint(ggcc(1))
	rstj("HL_Lei")=1
	rstj.update
	rstj.close
	set rstj=nothing
end if
rs.close
set rs=nothing

end if

'**********************
'加入或修改广告
'**********************

set rs=conn.execute("select * from "&jieducm&"_ggwei where HL_title='"&ggflei&"'")
ggflei=rs("HL_id")
rs.close
set rs=nothing

set rs=conn.execute("select * from "&jieducm&"_ggwei where HL_title='"&ggwei&"'")
ggwei=rs("HL_id")
rs.close
set rs=nothing



set rs = conn.execute("select Max(HL_ggID) From "&jieducm&"_Advertising where HL_ggwei="&cint(ggwei))
         Hl_Home=rs(0)
    if isnull(Hl_Home) then
         Hl_Home=1
	 else
	     Hl_Home=Hl_Home+1	 
    end if
    rs.close
   set rs=nothing    
	

	

set rsggtj=server.CreateObject("adodb.recordset")

if request("action")="HL_ggtj" then
rsggtj.open "select * from "&jieducm&"_Advertising",conn,1,3
rsggtj.addnew
rsggtj("HL_ggid")=HL_Home
rsggtj("HL_ggtime")=Now()
rsggtj("HL_deltime")=Now()
else
rsggtj.open "select * from "&jieducm&"_Advertising where id="&cint(request("id")),conn,1,3
end if
rsggtj("HL_ggtu")=request("HL_ggtu")
rsggtj("HL_ggwei")=cint(ggwei)
rsggtj("HL_gglei")=cint(ggflei)
rsggtj("HL_Title")=request("HL_Title")
rsggtj("HL_home")=request("HL_Urltj")
rsggtj("HL_ggImg")=request("Upfile")




rsggtj.update
rsggtj.close
set rsggtj=nothing
response.Redirect "HL_admin_gg.asp?action=ORder&id="&ggwei
end if


'***************************

sub ggweitile(id,title)
set rswei=conn.execute("select * from "&jieducm&"_ggwei where HL_id="&cint(id))
response.write rswei(""&title&"")
end sub


if request("action")="ORderb" or request("action")="ORderp" then

action=request("action")
HL_ggid=request("HL_ggid")
ggid=request("ggid")

if action="ORderb" then
if HL_ggid=0 then
response.Write "<script LANGUAGE='javascript'>alert('对不起，你没有选择要向下移动位数！');history.go(-1);</script>"
response.End()
end if
'1,3   1为现在的排位,3为要修改为的排位

'取得要更新数据的排名
'hl_ggid为要升到的位置
'NOtid为原来的位置
set rs=conn.execute("select * from "&jieducm&"_Advertising where id="&cint(ggid))
NOtid=rs("HL_ggid")
ggweiid=rs("HL_ggwei")
rs.close
set rs=nothing
response.write NOtid
conn.execute("update "&jieducm&"_Advertising set HL_ggid=HL_ggid-1 where HL_ggid>"&cint(NOtid)&" and HL_ggid<="&cint(hl_ggid)+cint(NOtid))
conn.execute("update "&jieducm&"_Advertising set HL_ggid=HL_ggid+"&HL_ggid&" where id="&cint(ggid))
else

if HL_ggid=0 then
response.Write "<script LANGUAGE='javascript'>alert('对不起，你没有选择要向上移动位数！');history.go(-1);</script>"
response.End()
end if
'4,2   3为现在的排位,1为要修改为的排位

'取得要更新数据的排名
'hl_ggid为要升到的位置  2
'NOtid为原来的位置   4
set rs=conn.execute("select * from "&jieducm&"_Advertising where id="&cint(ggid))
NOtid=rs("HL_ggid")
ggweiid=rs("HL_ggwei")
rs.close
set rs=nothing
response.write NOtid
conn.execute("update "&jieducm&"_Advertising set HL_ggid=HL_ggid+1 where HL_ggid<"&cint(NOtid)&" and HL_ggid>="&cint(NOtid)-cint(HL_ggid))
conn.execute("update "&jieducm&"_Advertising set HL_ggid=HL_ggid-"&HL_ggid&" where id="&cint(ggid))
end if
response.Redirect "HL_admin_gg.asp?action=ORder&id="&ggweiid
end if

if request("action")="del" then
  '删除所选会员
   if request("mode")="del" then
      if request("id")="" then
          response.Redirect "HL_admin_gg.asp"
          response.End()
      end if

     id=Split(request("id"),",")
     For i = LBound(id) To UBound(id)
       response.write id(i)&"<br>"

     
		  set rspai=conn.execute("select * from "&jieducm&"_Advertising where id="&cint(id(i)))
		  response.write rspai("HL_ggid")
		  
	   set rsxiu=conn.execute("update "&jieducm&"_Advertising set HL_ggid=HL_ggid-1 where HL_ggid>"&cint(rspai("HL_ggid"))&" and HL_ggwei="&cint(rspai("HL_ggwei")))
		'response.write id(i)
       set rs=conn.execute("delete  from "&jieducm&"_Advertising where id="&cint(id(i)))
	   		         
'        set rs=conn.execute("delete * from shops where UserID="&cint(id(i)))
'        set rs=conn.execute("delete * from UserDetails where UserID="&cint(id(i)))
        
		
        
     Next 


   end if
   '停用所选 会员
   if request("mode")="tingyong" then
   
      User_zt=request("User_zt")
	  
    if User_zt<>"" then 
	   response.write User_zt
	   response.write request("id")
	  ' response.End()
  
        set rs=conn.execute("update "&jieducm&"_Advertising set HL_ggzt="&User_zt&" where id="&cint(request("id")))
        set rs=nothing
        conn.close 
		response.Redirect "HL_admin_gg.asp"  
    end if
      if request("id")="" then
          response.Redirect "HL_admin_gg.asp"
          response.End()
      end if

     id=Split(request("id"),",")
     For i = LBound(id) To UBound(id)
       response.write id(i)&"<br>"

  
        set rs=conn.execute("update "&jieducm&"_Advertising set HL_ggzt=True where id="&cint(id(i)))
        set rs=nothing
        conn.close
     Next 
	 
	 
   end if
   '启用所选会员
   if request("mode")="qiyong" then
      if request("id")="" then
          response.Redirect "HL_admin_gg.asp"
          response.End()
      end if

     id=Split(request("id"),",")
     For i = LBound(id) To UBound(id)
       response.write id(i)&"<br>"

     
       set rs=conn.execute("update "&jieducm&"_Advertising set HL_ggzt=False where id="&cint(id(i)))
        set rs=nothing
        conn.close
     Next 
   end if
   if request("page")="" then
   page=1
   else
   page=request("page")
   end if
   response.Redirect "HL_admin_gg.asp?page="&request("page")
end if

%>