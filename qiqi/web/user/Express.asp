<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">				
<head>
<title>��Ա����-<%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<script language="javascript" src="../js/jquery-1.4.4.min.js"></script>
<style type="text/css">
.tb{
	background-color:#F4FCFF;
	border:1px solid #c4d5e0;
	margin-bottom:10px;
	height: 210px;
}
.Tips{
	background-color:#FEFFF0;
	padding: 8px;
	border: 1px solid #FFDEB9;
	display: block;
}
.orange,.orange a{color:#FF5500;}
.rough{font-weight:bold;}
.submiting{margin:10px; line-height:60px; height:60px; font-weight:bold; color:#0000FF; text-align:center; background:#F1F9FE url(../images/jieducm_shua.gif) 100px center no-repeat; border:1px solid #c7e5fa;}

.t_list td {border-bottom:1px solid #E5F5FF; line-height:24px; height:26px; clear:both; background-color:#FFF}
.t_list tr:hover td{background-color:#EBF5FF}
.t_list{
	margin-top:10px;
	width:710px;
	background-color:#eef9ff;
	border: 3px solid #10A2EF;
	margin-bottom: 10px;
	position: relative;
}
.t_list .title{
	height: 26px;
	font-weight: bold;
	font-size: 14px;
	display:block;
}
#retData{
	border: 1px solid #ACDAFF;
	background-color:#FFF;
	width:690px;
	margin: 10px;
}
#retData .Result{
	width: 690px;
	margin-left:10px;
	margin-top:10px;
}
#retData .num{
	line-height:22px;
	font-weight: bold;
	color:#F30;
	font-size: 14px;
}
#retData table, #retData .Result div{width: 690px; margin-bottom:10px;}
</style>
</head>

<body>

<!--#include file=../inc/jieducm_top.asp-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
				
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>
				<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>
				<div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;���ɿ�ݵ���</div>
                    <div class=pp8><strong>��ݵ����������ѯ����ѣ�</strong></div>

       <script language="javascript">
	 
	$(document).ready(function()
	{
		$("#btnSnap").click(function()
		{
		if (getRV("NumType")!="1"){
                    if ($("#expressno").val()==""){
                    alert('�������ݵ��ţ�');
                    $('#expressno').focus();return false; }
                }
		    $("#t_list").show();
			$("#retData").html('<div class="submiting">�˵��Ų�ѯ�У����Ե�.....</div>');
 
			var expressid = $("#expressid").val();
            var expressno = $("#expressno").val();
			 var numtype = getRV("NumType");	
			$.post("jieducm_express.asp",{com:expressid,nu:expressno,type:numtype,jieducm_rq:<%=stup_RndQueryNum%>,jieducm_dq:<%=stup_DefQueryNum%>,jieducm_vip:<%=vip%>},
				function(data)
				{		
					$("#retData").html(data);
				}
			);
 
			return false;
		});
	});
	        function getRV(b){var d="";var a=document.getElementsByName(b);for(var c=0;c<a.length;c++){if(a[c].checked){d=a[c].value}}return d}

	</script>
		<table width="100%" border="0" cellspacing="0" cellpadding="4" class="tb">
        <tr>
          <td align="right"></td>
          <td height="16" align="left"></td>
        </tr>
        <tr>
          <td width="40%" align="right"><strong>ѡ���ݹ�˾��</strong></td>
          <td width="60%" height="30" align="left">&nbsp;       <select name="expressid" id="expressid">
        <option selected="selected" value="yuantong">Բͨ���</option>
		<option  value="yunda">�ϴ���</option>
        <option value="zhongtong">��ͨ���</option>
        <option value="huitongkuaidi">��ͨ���</option>
		<option value="tiantian">������</option>
        <option value="zhaijisong">լ���Ϳ��</option>  
        </select> </td>
        </tr>
		 <tr>
          <td align="right"><strong>�������ɷ�ʽ��</strong></td>
          <td height="30" align="left">
          <label><input type="radio" name="NumType" id="NumType" value='1' checked="checked" onClick="$('#KuaiDiNum').hide();$('#expressno').val('');" />�������</label>&nbsp; 
          <label><input type="radio" name="NumType" id="NumType" value='2' onClick="$('#KuaiDiNum').show();" />�Զ�������</label>&nbsp; 
          <label title="����������Ŀ���˵���Ų�ѯ"><input type="radio" name="NumType" id="NumType" value='3' onClick="$('#KuaiDiNum').show();" />����ѯ�˵�</label></td>
        </tr>
        <tr id="KuaiDiNum" style="display:none;">
          <td align="right"><strong>��ݵ��ţ�</strong></td>
          <td height="30" align="left">&nbsp;<input name="expressno" type="text" id="expressno" value="" /></td>
        </tr>
        <tr>
          <td align="right">&nbsp;</td>
          <td height="50" align="left">&nbsp;<div class="txtButton">
		  <input type="image" name="Sumitbtn" id="btnSnap" src="../images/jieducm_ConfirmButton.gif" />		 
		  </div></td>
        </tr>			
		<tr>
			<td height="40" colspan="3" align="center" class="orange">ע�����ڲ�ѯ������һ��ϵͳ��Դ������ÿ�β�ѯ����10�롣</td>
		</tr>
      </table>
      

    <div class="t_list" id="t_list" style="display:none;">
        <span class="title">&nbsp; �����������ѯ�����</span>
        <div id="retData"></div>
    </div>
	
       <div class="Tips">
	  <p><span class="orange rough">������ɣ�</span>	  ����ʽ��ϵͳ������ѡ��ݹ�˾���Ź����������5�����š�<br />
	  <span class="orange rough">�Զ������ɣ�</span><br />
	    1.����ʽ�ʺ�׷�������ʵ�˵��ŵ��û�ʹ�ã�<br />
	    2.�����Խ��������е���ʵ��ݵ����������룬ϵͳ���Զ�ʶ�������ڽ����˵��ţ�<br />
	    3.��Ҳ����������ʵ�˵��ţ�ȥ������1-3λ���֣���ǰ��λ���֣�ϵͳ���Զ�ʶ�������ڽ����˵��ţ�������Ҫ���ݹ�˾ƥ�䣩��<br />
	    <span class="orange rough">���飺�����Ը��������˵����ջ���ַҪ����ַ��Ա��ջ���ַ��д�����ɵ��˵��ջ���Ϣ���Ա����˵���ѯֻ��ƥ�䵽����У������ʹ�ã� </span><br />
	    <span class="red rough">ע�⣺</span>Ϊ��ֹ���⣬��ͨ��Աÿ�����ʹ��<span class="orange rough"><%=stup_RndQueryNum%></span>��������ɺ�<span class="orange rough"><%=stup_DefQueryNum%></span>���Զ������ɣ�VIP��Ա�����ޣ�</p>
    </div>
      
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>
  <!--#include file=../inc/jieducm_bottom.asp-->
  <%call closeconn()%>
</body>
</html>