<%

dim ComeUrl,cUrl
ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
if ComeUrl="" then
	response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ��</font></p>"
	response.end
else
	cUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
	if mid(ComeUrl,len(cUrl)+1,1)=":" then
		cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
	end if
	cUrl=lcase(cUrl & request.ServerVariables("SCRIPT_NAME"))
	if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) then
		response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ����������ⲿ���ӵ�ַ���ʱ�ϵͳ��</font></p>"
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
 
 fileExt=lcase(file.FileExt)	'�õ����ļ���չ��������.

if file.filesize<100 then
 	response.write "<span style=""font-family: ����; font-size: 9pt"">����ѡ����Ҫ�ϴ����ļ�����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
	response.end
 end if
 if (filelx<>"swf") and (filelx<>"jpg") then 
 	response.write "<span style=""font-family: ����; font-size: 9pt"">���ļ����Ͳ����ϴ�����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
	response.end
 end if
 if filelx="swf" then
	if fileext<>"swf"  then
		response.write "<span style=""font-family: ����; font-size: 9pt"">ֻ���ϴ�swf��ʽ��Flash�ļ�����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
		response.end
	end if
 end if
 if filelx="jpg" then
	if fileext<>"gif" and fileext<>"jpg" and fileext<>"swf"  then
		response.write "<span style=""font-family: ����; font-size: 9pt"">ֻ���ϴ�jpg��gif��ʽ��ͼƬ����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
		response.end
     	end if
 end if
 if filelx="swf" then
	if file.filesize>(3000*1024) then
		response.write "<span style=""font-family: ����; font-size: 9pt"">���ֻ���ϴ� 3M ��Flash�ļ�����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
		response.end
	end if
 end if
 if filelx="jpg" then
	if file.filesize>(100*3024) then
		response.write "<span style=""font-family: ����; font-size: 9pt"">���ֻ���ϴ� 100K ��ͼƬ�ļ�����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
		response.end
	end if
 end if

 randomize
 ranNum=int(90000*rnd)+10000
 filename=filepath&"jieducm_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt

 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveToFile Server.mappath(FileName)
  'response.write file.FileName&"�����ϴ��ɹ�!����<br>"
  'response.write "���ļ�����"&FileName&"<br>"
  'response.write "���ļ����Ѹ��Ƶ������λ�ã��ɹرմ��ڣ�"
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
window.alert("�ļ��ϴ��ɹ�!�벻Ҫ�޸����ɵ����ӵ�ַ��");
window.close();
</script>

