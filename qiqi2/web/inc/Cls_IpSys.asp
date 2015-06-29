<%

function view_sys(vss,tt)
  dim vs,ver,strUserAgentArr,strTempUserInfo1,strTempUserInfo2,Mozilla,Agent,BcType,Browser,sSystem,strSystem,strBrowser
  on error resume next
  vs=trim(vss)
  strUserAgentArr=Split(vs,"; ")
  strTempUserInfo1=strUserAgentArr(1)
  if Instr(strTempUserInfo1,"MSIE")>0 then strTempUserInfo1=Replace(strTempUserInfo1,"MSIE","Internet Explorer")
  strTempUserInfo2=trim(Left(strUserAgentArr(2),Len(strUserAgentArr(2))-1))
  if InStr(vs,"Mozilla/4.0 (compatible;")>0 and strTempUserInfo2="Windows NT 5.0" then strTempUserInfo2="Windows 2000"
  Mozilla=vs
  Agent=Split(Mozilla,"; ")
  BcType=0
  If Instr(Agent(1),"U") Or Instr(Agent(1),"I") Then BcType=1
  If InStr(Agent(1),"MSIE") Then BcType=2
  select Case BcType
  Case 0
    Browser="其它"
    sSystem="其它"
  Case 1
    Ver=Mid(Agent(0),InStr(Agent(0),"/")+1)
    Ver=Mid(Ver,1,InStr(Ver," ")-1)
    Browser="Netscape" & Ver
    sSystem=Mid(Agent(0),InStr(Agent(0),"(")+1)
    sSystem=Replace(sSystem,"Windows","Win")
  case 2
    Browser=Agent(1):sSystem=Replace(Agent(2),")",""):sSystem=Replace(sSystem,"Windows","Win")
  End select
  strSystem=Replace(sSystem,"Win","Windows")
  if InStr(strSystem,"98")>0 and InStr(Mozilla,"Win 9x")>0 then strSystem=Replace(strSystem,"98","Me")
  strSystem=Replace(strSystem,"NT 5.2","2003")
  strSystem=Replace(strSystem,"NT5.2","2003")
  strSystem=Replace(strSystem,"NT 5.1","XP")
  strSystem=Replace(strSystem,"NT5.1","XP")
  strSystem=Replace(strSystem,"NT 5.0","2000")
  strSystem=Replace(strSystem,"NT5.0","2000")
  strSystem=Replace(strSystem,"NT 5.1","XP")
  strSystem=Replace(strSystem,"NT5.1","XP")
  strBrowser=Replace(Browser,"MSIE","Internet Explorer")
  set Browser=Nothing:set sSystem=Nothing
  if tt=1 then
    view_sys="操作系统：" & trim(strSystem) & "\n浏 览 器：" & trim(strBrowser)
  else
    view_sys="操作系统：" & trim(strSystem) &vbcrlf&"浏 览 器：" & trim(strBrowser)
  end if
  if err then err.clear:view_sys="未知的操作系统和浏览器"
End Function

function ip_true(tips)
  dim tip,iptemp,iptemp1,iptemp2
  tip=tips
  ip_true=false
  tip=trim(tip)
  iptemp=tip
  iptemp=replace(replace(iptemp,".",""),":","")
  iptemp1=tip
  iptemp1=len(tip)-len(replace(iptemp1,".",""))
  iptemp2=tip
  iptemp2=len(tip)-len(replace(iptemp2,":",""))
  if isnumeric(iptemp) and cint(iptemp1)=3 and (cint(iptemp2)=1 or cint(iptemp2)=0) then ip_true=true
End Function

function ip_ip(tips)
  dim ipn,tip
  tip=tips
  tip=trim(tip)
  if not ip_true(tip) then
    ip_ip="IP Error"
    exit function
  End If
  ipn=Instr(tip,":")
  if ipn>0 then
    ip_ip=left(tip,ipn-1)
    exit function
  End If
  ip_ip=tip
End Function

function ip_types(tips)
  dim ipn,tip2,wip,ip_type
  tip2=tips
  tip2=trim(tip2)
  ip_types="IP Error"
  if not ip_true(tip2) then exit function
  wip=YSvoid.Format_Mid_Num(10)
  if login_mode=format_power2(1,1) then wip=2
  select case wip
  case 0
    ip_types="*.*.*.*"
    ip_types="来访 I P："&ip_types
  case 1
    ipn=Instr(tip2,":")
    if ipn>0 then
      ip_types=left(tip2,ipn-1)
    else
      ip_types=tip2
    End If
    ip_type=split(ip_types,".")
    ip_types=ip_type(0)&"."&ip_type(1)&".*.*"
    erase ip_type
    ip_types="来访 I P："&ip_types
  case else
    ip_types=tip2
    ip_types="来访 I P："&tip2
  end select
End Function

function ip_port(pips)
  dim ipnn,iptemp,pip
  pip=pips
  pip=trim(pip)
  if ip_true(pip)="no" then
    ip_port="no"
    exit function
  End If
  ipnn=Instr(pip,":")
  if ipnn>0 then
    ip_port=right(pip,len(pip)-ipnn)
    exit function
  End If
  ip_port="yes"
End Function

function ip_address(sips)
  dim str1,str2,str3,str4,num,country,irs,sip
  sip=sips
  if not(isnumeric(left(sip,2))) then
    ip_address="未知"
    exit function
  End If
  if sip="127.0.0.1" then sip="192.168.0.1"
  str1=left(sip,instr(sip,".")-1)
  sip=mid(sip,instr(sip,".")+1)
  str2=left(sip,instr(sip,".")-1)
  sip=mid(sip,instr(sip,".")+1)
  str3=left(sip,instr(sip,".")-1)
  str4=mid(sip,instr(sip,".")+1)
  if not(isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0) then
	dim IPconn
    set IPconn=server.createobject("adodb.connection")
    set irs=server.createobject("adodb.recordset")
    IPconn.open IpConnstr
    num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
    sql="select Top 1 country from ip_address where ip1 <="&num&" and ip2 >="&num&""
    irs.open sql,IPconn,1,3
    if irs.eof then
      country="亚洲"
    else
      country=irs(0)
    End If
    irs.close
    set irs=nothing
    set IPconn=nothing
  End If
  ip_address=country
End Function

%>