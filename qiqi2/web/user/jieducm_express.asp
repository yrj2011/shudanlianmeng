<%
'�뱣����������Ϣ,�������������Ӱ������ٶ�!
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn http://www.jieducm.com
'�������ƽ̨��http://www.258shua.com
'��ʾƽ̨��ַ��http://www.258shua.net/
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'ע: �벻Ҫ���Ը��Ĵ˱�ʶ,����Ҫ˽�·���������,������Ϊ�Զ������ۺ����
'*****************************************************************

From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "����ʱ,��رձ�ҳ�沢���µ�½����!"
response.end
end If
Response.Charset="gb2312" 
Server.ScriptTimeout = 999999999
rq=request.form("jieducm_rq")
dq=request.form("jieducm_dq")
vip=request.form("jieducm_vip")
jieducm_type=request.form("type")
if vip<>1 then
userName=Request.Cookies("jieducm_name")("visiter") 
userName2=Request.Cookies("jieducm_name2")("visiter") 
Response.Cookies("jieducm_name")("visiter")=userName+1 
Response.Cookies("jieducm_name2")("visiter")=userName2+1 
Response.Cookies("jieducm_name").Expires= (now()+1) 
if (userName)>=(rq) and jieducm_type=1 then
Response.Write("��ͨ��Աÿ�����ʹ��"&rq&"���������")
response.End
end if
if (userName2)>=(dq) and jieducm_type=2 then
Response.Write("��ͨ��Աÿ�����ʹ��"&dq&"���Զ�������")
response.End
end if
end if
RefreshIntervalTime = 10 '��ֹˢ�µ�ʱ��������0��ʾ����ֹ
If Not IsEmpty(Session("QYCvisit")) and isnumeric(Session("QYCvisit")) and int(RefreshIntervalTime) > 0 Then
if (timer()-int(Session("QYCvisit")))*1000 < RefreshIntervalTime * 1000 then
Response.write ("<meta http-equiv=""refresh"" content="""& RefreshIntervalTime &""" />")
Response.write ("<b>ˢ��ʱ����Ϊ"& RefreshIntervalTime &"�룬���Ժ����ԣ�</b>")
Session("QYCvisit") = timer()
Response.end
end if
End If
Session("QYCvisit") = timer()

Dim Retrieval,NumArr,AppKey,SendURL,ResponseTxt,url,nu
AppKey = "39232c200e901622" 
com = Request.form("com")'��˾
jieducm_kdnu=request.form("nu")

if com="yuantong" then
jieducm_ji=0
jieducm_lennums=10
elseif com="yunda" then
jieducm_ji=1
jieducm_lennums=13
elseif com="zhongtong" then
jieducm_ji=2
jieducm_lennums=12
elseif com="huitongkuaidi" then
jieducm_ji=3
jieducm_lennums=12
elseif com="tiantian" then
jieducm_ji=4
jieducm_lennums=14
elseif com="zhaijisong" then
jieducm_ji=5
jieducm_lennums=10
end if
'������˵�������ģ�� (ֻ��Ҫ���뵥��ǰ��λ���ɣ����������ӣ�����ö̺ŷָ�)
jieducm_nu_NumArr = Array("17323,17713,17527,24913,24923,24951,25073,604984,61474,61489,61531,61546,61638,61691,61743,61801,61817,618170,61846,61894,7007,70079",_
"190013722,190013723,190013724,190013725,12002428,1200414888,12004369,12004339,1200437,15000283,16001051,19001166,19000936,19000946,19000956,19000966,190009605,19000971,80000148,8000014",_
"7170087,7170088,7170081,7170082,7170089,618181,618183,628015,628018,628020,628021,7010007,717000,7170007,718002,728002,757100,761003,762003,762004",_
"21005092,21005093,21005097,21005098,00001527,21005096,02000,21005094,21005099,280007",_
"0000112024,000001479,00001120245,00001120242,000001691,0000036212,0000037012,000008416,000001997,00001120247",_
"2990982,299438,540611,544039,576484,577609,647,6412005,642065,64297,6429765,642065,644087,644849,647975,671476,722274,7827788,782889,821115,9101187,911244,911930,912372,931646")

jieducm_num_rnd=Array("","9","99","999","9999","99999","999999","9999999","99999999","999999999","9999999999","99999999999","999999999999")

if jieducm_type="1" then

jieducm_nu = split(jieducm_nu_NumArr(jieducm_ji),",")
jieducm_dsk=Ubound(jieducm_nu)

for i=1 to 5
randomize
jieducm_mid=int(jieducm_dsk*rnd)
jieducm_nu_shen=int(jieducm_lennums)-int(len(jieducm_nu(jieducm_mid)))
jieducm_d=jieducm_num_rnd(jieducm_nu_shen)
jieducm_Cid=int(jieducm_d*rnd)
nu=jieducm_nu(jieducm_mid)&jieducm_Cid
SendURL ="http://api.kuaidi100.com/api?id="&AppKey&"&com="&com&"&nu="&nu&"&show=2&muti=1&order=asc"
ResponseTxt=GetHTTPPage(SendURL) '//��ȡԴ����ĺ�����������
response.Write("<font color=red><strong>�˵���:"&nu&"</strong></font>")'�����ѯ���
Response.Write ResponseTxt
next
elseif jieducm_type="2" then
jieducm_kdadd=jieducm_lennums-len(jieducm_kdnu)
jieducm_d=jieducm_num_rnd(jieducm_kdadd)
for jieducm_i=1 to 5
randomize
jieducm_Cid=int(jieducm_d*rnd)
jieducm_kdnu2=jieducm_kdnu&jieducm_Cid
SendURL ="http://api.kuaidi100.com/api?id="&AppKey&"&com="&com&"&nu="&jieducm_kdnu2&"&show=2&muti=1&order=asc"
ResponseTxt=GetHTTPPage(SendURL) '//��ȡԴ����ĺ�����������
response.Write("<font color=red><strong>�˵���:"&jieducm_kdnu2&"</strong></font>")'�����ѯ���
Response.Write ResponseTxt
next
elseif jieducm_type="3" then
SendURL ="http://api.kuaidi100.com/api?id="&AppKey&"&com="&com&"&nu="&jieducm_kdnu&"&show=2&muti=1&order=asc"
ResponseTxt=GetHTTPPage(SendURL) '//��ȡԴ����ĺ�����������
response.Write("<font color=red><strong>�˵���:"&jieducm_kdnu&"</strong></font>")'�����ѯ���
Response.Write ResponseTxt
end if
'���÷����������
Function GetHTTPPage(URL) 
    Dim objXML 
    Set objXML=CreateObject("MSXML2.SERVERXMLHTTP.3.0")  '����XMLHTTP��������Կռ��Ƿ�֧��XMLHTTP���������֧�֣����������������
	'Set objXML=Server.CreateObject("Microsoft.XMLHTTP") 
	'Set objXML=Server.CreateObject("MSXML2.XMLHTTP.4.0") 
	objXML.SetTimeouts 5000, 5000, 30000, 10000' 
    objXML.Open "GET",URL,False 'false��ʾ��ͬ���ķ�ʽ��ȡ��ҳ���룬�˽�ʲô��ͬ����ʲô���첽��
    objXML.Send() '����
	If objXML.Readystate<>4 Then
		Exit Function 
	End If
	GetHTTPPage=BytesToBstr(objXML.ResponseBody)'������Ϣ��ͬʱ�ú���������룬�������Ҫת����ѡ�� 
    Set objXML=Nothing'�ر� 
	If Err.number<>0 Then 
		Response.Write "<p align='center'><font color='red'><b>��������ȡ�ļ����ݳ������Ժ����ԣ�</b></font></p>" 
		Err.Clear
	End If
End Function

'ҳ�����ת��
Function BytesToBstr(body) 
    Dim objstream 
    Set objstream = Server.CreateObject("Adodb.Stream") '//����adodb.stream���
    objstream.Type = 1 
    objstream.Mode =3 
    objstream.Open 
    objstream.Write body 
    objstream.Position = 0 
    objstream.Type = 2 
    objstream.Charset = "UTF-8" 'ת��ԭ��Ĭ�ϵı���ת����GB2312���룬����ֱ����XMLHTTP�����������ַ�����ҳ�õ��Ľ������� 
    BytesToBstr = objstream.ReadText 
    objstream.Close 
    Set objstream = Nothing 
End Function
%>