<%
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	if totalnumber>0 then
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
	strTemp=strTemp & "�� <font color=blue><b>" & totalnumber & "</b></font> " & strUnit & "&nbsp;&nbsp;&nbsp;"
	strUrl=JoinChar(sfilename)
	if PageNo<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "PageNo=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "PageNo=" & (PageNo-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-PageNo<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "PageNo=" & (PageNo+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "PageNo=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & PageNo & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "PageNo=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(PageNo)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	response.write strTemp
	end if
end sub

%>