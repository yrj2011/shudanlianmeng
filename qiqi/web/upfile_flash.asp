<%

dim ComeUrl,cUrl
ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
if ComeUrl="" then
	response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统。</font></p>"
	response.end
else
	cUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
	if mid(ComeUrl,len(cUrl)+1,1)=":" then
		cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
	end if
	cUrl=lcase(cUrl & request.ServerVariables("SCRIPT_NAME"))
	if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) then
		response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许从外部链接地址访问本系统。</font></p>"
		response.end
	end if
end if
				
if session("AdminName")<>"" then
else
response.Redirect("../login.asp")
response.End()
end if
%>
<!--#include file="upload_wj110.inc"-->
<%
set upload=new upload_file
if upload.form("act")="uploadfile" then
	filepath=trim(upload.form("filepath"))
	filelx=trim(upload.form("filelx"))
	i=0
	for each formName in upload.File
	set file=upload.File(formName)
 
 fileExt=lcase(file.FileExt)	'得到的文件扩展名不含有.

if file.filesize<100 then
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">请先选择你要上传的文件！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
	response.end
 end if
 if (filelx<>"swf") and (filelx<>"jpg") then 
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">该文件类型不能上传！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
	response.end
 end if
 if filelx="swf" then
	if fileext<>"swf"  then
		response.write "<span style=""font-family: 宋体; font-size: 9pt"">只能上传swf格式的Flash文件！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
		response.end
	end if
 end if
 if filelx="jpg" then
	if fileext<>"gif" and fileext<>"jpg" and fileext<>"swf"  then
		response.write "<span style=""font-family: 宋体; font-size: 9pt"">只能上传jpg或gif格式的图片！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
		response.end
     	end if
 end if
 if filelx="swf" then
	if file.filesize>(3000*1024) then
		response.write "<span style=""font-family: 宋体; font-size: 9pt"">最大只能上传 3M 的Flash文件！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
		response.end
	end if
 end if
 if filelx="jpg" then
	if file.filesize>(100*3024) then
		response.write "<span style=""font-family: 宋体; font-size: 9pt"">最大只能上传 100K 的图片文件！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
		response.end
	end if
 end if

 randomize
 ranNum=int(90000*rnd)+10000
 filename=filepath&"jieducm_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt

 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveToFile Server.mappath(FileName)
  'response.write file.FileName&"　　上传成功!　　<br>"
  'response.write "新文件名："&FileName&"<br>"
  'response.write "新文件名已复制到所需的位置，可关闭窗口！"
  if filelx="swf" then
      response.write "<script>window.opener.document."&upload.form("FormName")&".size.value='"&int(file.FileSize/1024)&" K'</script>"
  end if
  response.write "<script>window.opener.document."&upload.form("FormName")&"."&upload.form("EditName")&".value='"&FileName&"'</script>"

 end if
 set file=nothing
	next
	set upload=nothing
end if
%>
<script language="javascript">
window.alert("文件上传成功!请不要修改生成的链接地址！");
window.close();
</script>

