function Show_Style()
{
	document.all["c0"].innerHTML="<table width=318 height=39 border=0 cellpadding=0 cellspacing=1 bgcolor=#0099CC><tr><td width=316 bgcolor=#AECFE3><div align=center><span id=c0>*4-15 个字符 (仅限大小写字母,数字,下划线等)。</span></div></td></tr></table>";
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
		   ||str.charAt(loop_index) == '　'
		   ||str.charAt(loop_index) == '１'
		   ||str.charAt(loop_index) == '２'
		   ||str.charAt(loop_index) == '３'
		   ||str.charAt(loop_index) == '４'		   
		   ||str.charAt(loop_index) == '５'
		   ||str.charAt(loop_index) == '６'		   
		   ||str.charAt(loop_index) == '７'
		   ||str.charAt(loop_index) == '８' 
		   ||str.charAt(loop_index) == '９'
		   ||str.charAt(loop_index) == '０' 
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
		if(a==""){msg="用户名不可以为空";document.all["c0"].innerHTML="* "+msg;b=false;}
		else{
			if(a.length < 4){msg="帐号长度少于4位";document.all["c0"].innerHTML="* "+msg;b=false;}
			else{
				if(checkC(a)){
					if(checksql(a)){document.all["tname"].src="jieducm_checkname.asp?n="+a;}
					else{msg="含有非法语法字符!!";document.all["c0"].innerHTML="*"+msg;b=false;}
								}
				else
				{msg="含有非法字符!";document.all["c0"].innerHTML="*"+msg;b=false;}
				}
			}
	return b;
		}
