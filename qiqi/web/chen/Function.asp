<%

function ReplaceBadChar(strChar)
	if strChar="" then
		ReplaceBadChar=""
	else
                ReplaceBadChar=replace(replace(replace(replace(replace(replace(replace(strChar,"'",""),"*",""),"?",""),"(",""),")",""),"<",""),".","")
	end if
end function
AdminId=request.Cookies("AdminId")
AdminName=ReplaceBadChar(request.Cookies("AdminName"))
AdminPass=ReplaceBadChar(request.Cookies("AdminPass"))
sql="select top 1 * from ["&jieducm&"_dj586_Admin]"
set RsAdmin=server.createobject("adodb.recordset")
RsAdmin.open sql,conn,1,1
IF RsAdmin.eof Then
errmsg=errmsg+ "<BR>"+"<li>��û�н��뱾ҳ���Ȩ��!���β����ѱ���¼!<li>��������û�е�½���߲�����ʹ�õ�ǰ���ܵ�Ȩ��!����ϵ����Ա.<li>��ҳ��Ϊ[<font color=red>����Ա</font>]ר��,���ȵ�½�����."
call Error_Msg(Errmsg) 
response.end
else
dj586_Com_AdminID=RsAdmin("ID")
dj586_Com_AdminName=RsAdmin("AdminName")
dj586_Com_AdminFlag=RsAdmin("Flag")
dj586_Com_AdminAa=RsAdmin("Aa")
dj586_Com_AdminBb=RsAdmin("Bb")
dj586_Com_AdminCc=RsAdmin("Cc")
dj586_Com_AdminDd=RsAdmin("Dd")
dj586_Com_AdminEe=RsAdmin("Ee")
dj586_Com_AdminFf=RsAdmin("Ff")
dj586_Com_AdminGg=RsAdmin("Gg")
dj586_Com_AdminHh=RsAdmin("Hh")
dj586_Com_AdminIi=RsAdmin("Ii")
dj586_Com_AdminJj=RsAdmin("Jj")
dj586_Com_AdminKk=RsAdmin("Kk")
RsAdmin.close
set RsAdmin=nothing
End IF
%>