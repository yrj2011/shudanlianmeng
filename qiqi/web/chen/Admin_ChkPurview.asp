<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
dim AdminName,AdminPurview,PurviewPassed
dim AdminPurview_Article,jieducm_addrssweb,AdminPurview_Soft,AdminPurview_Photo,AdminPurview_Guest,AdminPurview_Others
dim AdminPurview_class1,AdminPurview_class2,AdminPurview_class3,AdminPurview_class4,AdminPurview_class5,AdminPurview_class6
dim AdminPurview_class7,AdminPurview_class8,AdminPurview_class9,AdminPurview_class10,AdminPurview_class11,AdminPurview_class12
dim AdminPurview_class13,AdminPurview_class14,AdminPurview_class15,AdminPurview_class16,AdminPurview_class17,AdminPurview_class18
dim AdminPurview_class19,AdminPurview_class20,AdminPurview_class21,AdminPurview_class22,AdminPurview_class31,AdminPurview_class32
dim rsGetAdmin,sqlGetAdmin
dim ComeUrl,cUrl
ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))


AdminName=replace(session("AdminName"),"'","")
if AdminName="" then
	call CloseConn()
	response.redirect "Admin_login.asp"
end if
sqlGetAdmin="select * from "&jieducm&"_stup"
set rsGetAdmin=nt2003.execute(sqlGetAdmin)
if rsGetAdmin.bof and rsGetAdmin.eof then
	rsGetAdmin.close
	set rsGetAdmin=nothing
	'call CloseConn()
	response.redirect "Admin_login.asp"
else
AdminPurview_class31=rsGetAdmin("jieducm_webad")
AdminPurview_class32=rsGetAdmin("jieducm_adweb")
end if
sqlGetAdmin="select * from "&jieducm&"_Admin where UserName='" & AdminName & "'"
set rsGetAdmin=nt2003.execute(sqlGetAdmin)
if rsGetAdmin.bof and rsGetAdmin.eof then
	rsGetAdmin.close
	set rsGetAdmin=nothing
	'call CloseConn()
	response.redirect "Admin_login.asp"
end if
AdminPurview=rsGetAdmin("Purview")
AdminPurview_Article=rsGetAdmin("AdminPurview_Article")
AdminPurview_Guest=rsGetAdmin("AdminPurview_Guest")
AdminPurview_Others=rsGetAdmin("AdminPurview_Others")
AdminPurview_class1=rsGetAdmin("class1")
AdminPurview_class2=rsGetAdmin("class2")
AdminPurview_class3=rsGetAdmin("class3")
AdminPurview_class4=rsGetAdmin("class4")
AdminPurview_class5=rsGetAdmin("class5")
AdminPurview_class6=rsGetAdmin("class6")
AdminPurview_class7=rsGetAdmin("class7")
AdminPurview_class8=rsGetAdmin("class8")
AdminPurview_class9=rsGetAdmin("class9")
AdminPurview_class10=rsGetAdmin("class10")
AdminPurview_class11=rsGetAdmin("class11")
AdminPurview_class12=rsGetAdmin("class12")
AdminPurview_class13=rsGetAdmin("class13")
AdminPurview_class14=rsGetAdmin("class14")
AdminPurview_class15=rsGetAdmin("class15")
AdminPurview_class16=rsGetAdmin("class16")
AdminPurview_class17=rsGetAdmin("class17")
AdminPurview_class18=rsGetAdmin("class18")
AdminPurview_class19=rsGetAdmin("class19")
AdminPurview_class20=rsGetAdmin("class20")
AdminPurview_class21=rsGetAdmin("class21")
AdminPurview_class22=rsGetAdmin("class22")
rsGetAdmin.close
set rsGetAdmin=nothing
PurviewPassed=True
if PurviewLevel>0 then
	if AdminPurview>PurviewLevel then
		PurviewPassed=False
	else
		if AdminPurview=2 then
			select case CheckChannelID
				case 0        '�����������
					PurviewPassed=CheckPurview(AdminPurview_Others,PurviewLevel_Others)
				case 2        '����Ƶ��
					if AdminPurview_Article>PurviewLevel_Article then
						PurviewPassed=False
					end if
				case 5       '���԰�
					if AdminType=True then
						PurviewPassed=CheckPurview(AdminPurview_Guest,PurviewLevel_Guest)
					else
						PurviewPassed=True
					end if
			end select
		end if
	end if
end if
if PurviewPassed=False then
	response.write "<br><p align=center><font color='red'>�Բ�����û�д��������Ȩ�ޡ�</font></p>"
	response.end
end if

function CheckPurview(AllPurviews,strPurview)
	if isNull(AllPurviews) or AllPurviews="" or strPurview="" then
		CheckPurview=False
		exit function
	end if
	CheckPurview=False
	if instr(AllPurviews,",")>0 then
		dim arrPurviews,i
		arrPurviews=split(AllPurviews,",")
		for i=0 to ubound(arrPurviews)
			if trim(arrPurviews(i))=strPurview then
				CheckPurview=True
				exit for
			end if
		next
	else
		if AllPurviews=strPurview then
			CheckPurview=True
		end if
	end if
end function

function CheckClassMaster(AllMaster,MasterName)
	if isNull(AllMaster) or AllMaster="" or MasterName="" then
		CheckClassMaster=False
		exit function
	end if
	CheckClassMaster=False
	if instr(AllMaster,"|")>0 then
		dim arrMaster,i
		arrMaster=split(AllMaster,"|")
		for i=0 to ubound(arrMaster)
			if trim(arrMaster(i))=MasterName then
				CheckClassMaster=True
				exit for
			end if
		next
	else
		if AllMaster=MasterName then
			CheckClassMaster=True
		end if
	end if
end function
%>
<!--#include file="Admin_PopMenu.asp"-->
