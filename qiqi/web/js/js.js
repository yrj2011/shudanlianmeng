function Show_Style()
{
	document.all["c0"].innerHTML="<table width=318 height=39 border=0 cellpadding=0 cellspacing=1 bgcolor=#0099CC><tr><td width=316 bgcolor=#AECFE3><div align=center><span id=c0>*4-15 ���ַ� (���޴�Сд��ĸ,����,�»��ߵ�)��</span></div></td></tr></table>";
}
function  checkC(str) {
       for(var loop_index=0; loop_index<str.length; loop_index++)  
       { 
	    
         if(str.charAt(loop_index) == '~'   
           ||str.charAt(loop_index) == '!'  
           ||str.charAt(loop_index) == '@'  
           ||str.charAt(loop_index) == '#'  
           ||str.charAt(loop_index) == '$'  
           ||str.charAt(loop_index) == '%'  
           ||str.charAt(loop_index) == '^'  
           ||str.charAt(loop_index) == '&'  
           ||str.charAt(loop_index) == '*'  
           ||str.charAt(loop_index) == '('  
           ||str.charAt(loop_index) == ')'  
           ||str.charAt(loop_index) == '+'  
           ||str.charAt(loop_index) == '{'  
           ||str.charAt(loop_index) == '}'  
           ||str.charAt(loop_index) == '|'  
           ||str.charAt(loop_index) == ':'  
           ||str.charAt(loop_index) == '"'  
           ||str.charAt(loop_index) == '<'  
           ||str.charAt(loop_index) == '>'  
           ||str.charAt(loop_index) == '?'  
           ||str.charAt(loop_index) == '`'  
           ||str.charAt(loop_index) == '='  
           ||str.charAt(loop_index) == '['  
           ||str.charAt(loop_index) == ']'  
           ||str.charAt(loop_index) == '\\'  
           ||str.charAt(loop_index) == ';'  
           ||str.charAt(loop_index) == '\''  
           ||str.charAt(loop_index) == ','  
           ||str.charAt(loop_index) == '.'  
           ||str.charAt(loop_index) == '-'
		   ||str.charAt(loop_index) == ' '
		   ||str.charAt(loop_index) == '��'
		   ||str.charAt(loop_index) == '��'
		   ||str.charAt(loop_index) == '��'
		   ||str.charAt(loop_index) == '��'
		   ||str.charAt(loop_index) == '��'		   
		   ||str.charAt(loop_index) == '��'
		   ||str.charAt(loop_index) == '��'		   
		   ||str.charAt(loop_index) == '��'
		   ||str.charAt(loop_index) == '��' 
		   ||str.charAt(loop_index) == '��'
		   ||str.charAt(loop_index) == '��' 
		   ||str.charAt(loop_index) == '/') 
          {  
            //alert("~,,,!,@,#,$,%,^,&,*,+,`,\',\",:,(,),[,],{,},<,>,|,\\ and / are illegal. Please re-input."); 
            //alert(errmsg);
            return false;  
      }  
         }//end of for(loop_index)  
      return true;
   }
function  checksql(str) {
	//alert(str);
	return true;
	var ary=['select','update','insert','exec','and','drop'];
	for(var i=0;i<ary.length;i++)
	{if(str.indexOf(ary[i])){return false;
	}
	}
	return true;
	}
//	function dd(a)
//	{
//		var a=a.value;
//		document.all["tname"].src="testname.asp?n="+a;	
//	}
function name_chk(a){
	var b=true; 
	theobj=a;
	var a=a.value;
		if(a==""){msg="�û���������Ϊ��";document.all["c0"].innerHTML="* "+msg;b=false;}
		else{
			if(a.length < 4){msg="�ʺų�������4λ";document.all["c0"].innerHTML="* "+msg;b=false;}
			else{
				if(checkC(a)){
					if(checksql(a)){document.all["tname"].src="jieducm_checkname.asp?n="+a;}
					else{msg="���зǷ��﷨�ַ�!!";document.all["c0"].innerHTML="*"+msg;b=false;}
								}
				else
				{msg="���зǷ��ַ�!";document.all["c0"].innerHTML="*"+msg;b=false;}
				}
			}
	return b;
		}
