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
errmsg=errmsg+ "<BR>"+"<li>您没有进入本页面的权限!本次操作已被记录!<li>可能您还没有登陆或者不具有使用当前功能的权限!请联系管理员.<li>本页面为[<font color=red>管理员</font>]专用,请先登陆后进入."
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