<%@language=vbscript codepage=936 %>
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
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=3
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->

<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim ArticleID,Action,sqlProperty,rsProperty,FoundErr,ErrMsg,PurviewChecked
dim ClassID,tClass,ClassName,RootID,ParentID,Depth,ParentPath,Child,ChildID,tID,tChild,ClassMaster
dim sqlmessage,adname ''�˸�����Ӵ���
ArticleID=trim(request("ArticleID"))
Action=Trim(Request("Action"))
adname=session("AdminName")
FoundErr=False
if ArticleId="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>����ѡ�����£�</li>"
end if
if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
end if
if FoundErr=False then
	call SetProperty()
end if
if FoundErr=False then
	'call CloseConn()
        if Action="nopassed" then
	response.Redirect "Admin_ArticleCheck.asp"
	else
	response.Redirect ComeUrl
	end if
else
	'call CloseConn()
	call WriteErrMsg()
end if

sub SetProperty()
	if instr(ArticleID,",")>0 then
		ArticleID=replace(ArticleID," ","")
		sqlProperty="select * from "&jieducm&"_Article where ArticleID in (" & ArticleID & ")"
	else
		ArticleID=Clng(ArticleID)
		sqlProperty="select * from "&jieducm&"_article where ArticleID=" & ArticleID
	end if
	Set rsProperty= Server.CreateObject("ADODB.Recordset")
	rsProperty.open sqlProperty,conn,1,3
	do while not rsProperty.eof
		PurviewChecked=False
		ClassID=rsProperty("ClassID")
		if AdminPurview=1 or AdminPurview_Article<=2 then
			PurviewChecked=True
		else
			if Action="SetPassed" or Action="CancelPassed" then
				set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassChecker From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
			else
				set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassMaster From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
			end if
			if tClass.bof and tClass.eof then
				founderr=True
				ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ</li>"
			else
				ClassName=tClass(0)
				RootID=tClass(1)
				ParentID=tClass(2)
				Depth=tClass(3)
				ParentPath=tClass(4)
				Child=tClass(5)
				ClassMaster=tClass(6)
				PurviewChecked=CheckClassMaster(tClass(6),AdminName)
				if PurviewChecked=False and ParentID>0 then
					if Action="SetPassed" or Action="CancelPassed" then
						set tClass=nt2003.execute("select ClassChecker from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
					else
						set tClass=nt2003.execute("select ClassMaster from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
					end if
					do while not tClass.eof
						PurviewChecked=CheckClassMaster(tClass(0),AdminName)
						if PurviewChecked=True then exit do
						tClass.movenext
					loop
				end if
			end if
		end if
		if PurviewChecked=False then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�� " & rsProperty("ArticleID") & " û�в���Ȩ�ޣ�</li>"
		else
			select case Action
			case "SetOnTop"
				rsProperty("OnTop") = True
				    '����ӵ� ���ö��ӣ�  һ֦÷��ӿ�ʼ
		          	'if nt2003.site_setting(37)=1 then
			        '   Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "+" & nt2003.site_setting(42) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'����
			case "CancelOnTop"
				rsProperty("OnTop") = False
				    '����ӵ� ��ȡ���ö��ۣ�  һ֦÷��ӿ�ʼ
		          	'if nt2003.site_setting(37)=1 then
			        '   Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(42) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		          '  end if
					'����
			case "SetElite"
				rsProperty("Elite") = True
				    '����ӵ� �����Ƽ��ӣ�  һ֦÷��ӿ�ʼ
		          '	if nt2003.site_setting(37)=1 then
			       '    Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "+" & nt2003.site_setting(41) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'����
			case "CancelElite"
				rsProperty("Elite") = False
				    '����ӵ� ��ȡ���Ƽ��ۣ�  һ֦÷��ӿ�ʼ
		          	'if nt2003.site_setting(37)=1 then
			       '    Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(41) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'����
			case "SetPassed"
				rsProperty("Passed") =True
				'Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleChecked & "=" & db_User_ArticleChecked & "+1 where " & db_User_Name & "='" & rsProperty("Editor") & "'")
				    '����ӵ� ��ͨ����˼ӣ�  һ֦÷��ӿ�ʼ
		          	'if nt2003.site_setting(37)=1 then
			         '  Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "+" & nt2003.site_setting(38) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'����
				'�˸�����Ӵ��뿪ʼ
				'SmsBody="Article_show.asp?ArticleID="&rsProperty("ArticleID")
 				'SmsBody="[url="&SmsBody&"]�������鿴���µľ�������[/url]"
				'sqlmessage="insert into "&db_User_Message&" (incept,sender,title,content,sendtime,flag,issend) values ('" & rsProperty("Editor") & "','��վ����Ա"&adname&"','����������¡�"&rsProperty("Title")&"���Ѿ�ͨ�����','"&SmsBody&"',Now(),0,1)"
				'Conn_User.execute(sqlmessage)
				nt2003.delcahe("NewArticle")
			case "nopassed"
				'rsProperty("nopass") =True
				'rsProperty("nopassno") =rsProperty("nopassno")+1
				'rsProperty("nopasstxt") =nopasstxt
				'SmsBody=""&rsProperty("Title")
            	'SmsBody="���ڣ�[color=bb3333]"&nopasstxt&"[/color]������������¡�"&SmsBody&"��δ�����ͨ�����뵽��̨����ѡ�񱻹���Ա�˸�����£��޸ĺ�������Ͷ�塣"
				'sqlmessage="insert into "&db_User_Message&" (incept,sender,title,content,sendtime,flag,issend) values ('" & rsProperty("Editor") & "','��վ����Ա"&adname&"','����������¡�"&rsProperty("Title")&"����������ԭ���ѱ�����Ա��"&rsProperty("nopassno")&"���˻ء�','"&SmsBody&"',Now(),0,1)"
				'Conn_User.execute(sqlmessage)
					'����ӵ� ���˸�ۣ�  һ֦÷��ӿ�ʼ
		          	'if nt2003.site_setting(37)=1 then
			        '  Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(39) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		            'end if
					'����
			case "deleted"
					'����ӵ� ��ɾ���ۣ�  һ֦÷��ӿ�ʼ
		          '	if nt2003.site_setting(37)=1 then
			         '  Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(40) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		          '  end if
					'����
				'SmsBody=""&rsProperty("Title") 	
				'SmsBody="����������¡�"&SmsBody&"��δ�����ͨ�����Ѿ�ɾ������������վȫ�������Ա������ʾʮ���ź���"
				'sqlmessage="insert into "&db_User_Message&" (incept,sender,title,content,sendtime,flag,issend) values ('" & rsProperty("Editor") & "','��վ����Ա"&adname&"','����������¡�"&rsProperty("Title")&"��δͨ����ˣ��Ѿ�ɾ��','"&SmsBody&"',Now(),0,1)"
				'Conn_User.execute(sqlmessage)
				conn.execute("delete from "&jieducm&"_Article where ArticleID in (" & ArticleID & ")")
				conn.execute("delete from "&jieducm&"_ArticleComment where ArticleID in (" & ArticleID & ")")
				'�˸�����Ӵ������
			case "CancelPassed"
				rsProPerty("Passed") =False
				'Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleChecked & "=" & db_User_ArticleChecked & "-1 where " & db_User_Name & "='" & rsProperty("Editor") & "'")
					'�����ӵ� ��ȡ����˿ۣ�  һ֦÷��ӿ�ʼ
		          	'if nt2003.site_setting(37)=1 then
			          ' Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(38) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'����
			end select
			rsProperty.update
		end if
		rsProperty.movenext
	loop
	rsProperty.close
	set rsProperty=nothing
end sub
%>