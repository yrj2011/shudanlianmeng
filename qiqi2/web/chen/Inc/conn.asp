<%
dim nt2003,template,SqlNowString,Conn,jieducm,admin_czmpass
Const IsSqlDataBase = 1
Const IsDeBug = 1
Set nt2003 = New System_Cls
'Set template = New cls_templates
Dim Fy_Get, Fy_In, Fy_Inf, Fy_Xh,FY_In2,Fy_Inf2
'自定义需要过滤的字串,用 "|" 分隔
Fy_In = "'|and|exec|insert|select|delete|update|count|chr|mid|master|truncate|char|declare"
Fy_In2 = "UNION|'|drop|and|exec|insert|select|delete|update|count|chr|mid|master|truncate|char|declare"
Fy_Inf = Split(Fy_In, "|")
Fy_Inf2 = Split(Fy_In2, "|")
If Request.QueryString <> "" Then
    For Each Fy_Get In Request.QueryString
        For Fy_Xh = 0 To UBound(Fy_Inf2)
            If InStr(LCase(Request.QueryString(Fy_Get)), Fy_Inf2(Fy_Xh)) <> 0 Then
                Response.Write "<Script Language=JavaScript>alert('请不要在参数中包含非法字符尝试注入！');</Script>"
                Fy_Get = ""
                Fy_In = ""
                Fy_Inf = ""
                Fy_Xh = ""
                Response.End
            End If
        Next
    Next
End If



if IsSqlDataBase = 1 then
SqlNowString = "GetDate()"
else
SqlNowString = "Now()"
end if

	jieducm="jieducm_gong"
	admin_czmpass="49ba59abbe56e057"'后台操作码
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open "Driver={Sql Server};Server=122.10.92.222,1433;UID=liuzhong1177;PWD=liuzhong1177a;Database=liuzhong1177"
'Conn.Open "provider=sqloledb;data source=122.10.92.222,1433;User ID=liuzhong1177;pwd=liuzhong1177a;Initial Catalog=liuzhong1177"

sub CloseConn()
On Error Resume Next
	Conn.close
	set Conn=nothing
end sub
Function HtmlEncode(Content)
  Content = Replace(Content, ">", "&gt;") 
  Content = Replace(Content, "<", "&lt;")
  Content = Replace(Content, "'", "") 
  HtmlEncode = content 
End Function

Function HtmlEncode2(Content)
  Content = Replace(Content, ">", "&gt;") 
  Content = Replace(Content, "<", "&lt;")
  Content = Replace(Content, " ", "&nbsp;")
  Content = Replace(Content, "'", "")
  Content = Replace(Content, vbcrlf,"<br>") 
  HtmlEncode2 = content 
End Function

function isChkInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isChkInteger=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isChkInteger=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isChkInteger=false 
              exit function
           end if
       next
       isChkInteger=true
       if err.number<>0 then err.clear
end function

Function PageSplit(currentpage,totalpage,pagename)                            
if currentpage mod 10 = 0 then                            
Sp = currentpage \ 10                            
else                            
Sp = currentpage \ 10 + 1                            
end if                            
Pagestart = (Sp-1)*10+1                            
Pageend = Sp*10                            
strSplit = "<a href="&pagename&"?pageid=1><font face=webdings title=第一页>9</font></a>&nbsp;"                            
if Sp > 1 then strSplit = strSplit & "<a href="&pagename&"?pageid="&Pagestart-10&"><font face=webdings title=前十页>7</font></a>&nbsp;"                            
for j=PageStart to Pageend                            
if j > totalpage then exit for                            
if j <> currentpage then                            
strSplit = strSplit & "<a href="&pagename&"?pageid="&j&">["&j&"]</a>&nbsp;"			                            
else                            
strSplit = strSplit & "<font color=red>["&j&"]</font>&nbsp;"                            
end if                            
next                            
if Sp*10  < totalpage then strSplit = strSplit & "<a href="&pagename&"?pageid="&Pagestart+10&"><font face=webdings title=后十页>8</font></a>"                             
strSplit = strSplit & "<a href="&pagename&"?pageid="&totalpage&" ><font face=webdings title=""最后一页"">:</font></a>"                            
PageSplit = strSplit                            
End Function

%>