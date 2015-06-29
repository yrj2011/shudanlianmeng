
<%
'Content--->用户输入的需要转换的字符串
'Number1--->没转换时切取的字符数量
'Number2--->转换后切取的字符数量 
Function HtmlToStr(Content,Number1,Number2)
  dim Str,Html,Num1,Num2,i
  Html=trim(Content) 
  Num1=Number1
  Num2=Number2
  if Num1>0 then 
     Html=left(Html,Num1)
  else
     Num1=len(Html)
  end if
  Html=replace(Html,"&nbsp;","")
  for i=1 to Num1
    if mid(Html,i,1)="<" then
    while (mid(Html,i,1)<>">") and (i<=Num1)
       i=i+1
    wend
 else
   select case mid(Html,i,1)
          case chr(34) : Str=Str
          case chr(13) : Str=Str
          case chr(9)  : Str=Str
          case chr(32) : Str=Str
          case else    : Str=Str+mid(Html,i,1)
      end select
 end if 
  next
  if Num2>0 then Str=left(Str,Num2)
  HtmlToStr=Str
End Function
%>