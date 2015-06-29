<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">				
<head>
<title>会员中心-<%=stupname%> -<%=stuptitle%></title>
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
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;生成快递单号</div>
                    <div class=pp8><strong>快递单号生成与查询（免费）</strong></div>

       <script language="javascript">
	 
	$(document).ready(function()
	{
		$("#btnSnap").click(function()
		{
		if (getRV("NumType")!="1"){
                    if ($("#expressno").val()==""){
                    alert('请输入快递单号！');
                    $('#expressno').focus();return false; }
                }
		    $("#t_list").show();
			$("#retData").html('<div class="submiting">运单号查询中，请稍等.....</div>');
 
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
          <td width="40%" align="right"><strong>选择快递公司：</strong></td>
          <td width="60%" height="30" align="left">&nbsp;       <select name="expressid" id="expressid">
        <option selected="selected" value="yuantong">圆通快递</option>
		<option  value="yunda">韵达快递</option>
        <option value="zhongtong">中通快递</option>
        <option value="huitongkuaidi">汇通快递</option>
		<option value="tiantian">天天快递</option>
        <option value="zhaijisong">宅急送快递</option>  
        </select> </td>
        </tr>
		 <tr>
          <td align="right"><strong>单号生成方式：</strong></td>
          <td height="30" align="left">
          <label><input type="radio" name="NumType" id="NumType" value='1' checked="checked" onClick="$('#KuaiDiNum').hide();$('#expressno').val('');" />随机生成</label>&nbsp; 
          <label><input type="radio" name="NumType" id="NumType" value='2' onClick="$('#KuaiDiNum').show();" />自定义生成</label>&nbsp; 
          <label title="根据您输入的快递运单编号查询"><input type="radio" name="NumType" id="NumType" value='3' onClick="$('#KuaiDiNum').show();" />仅查询运单</label></td>
        </tr>
        <tr id="KuaiDiNum" style="display:none;">
          <td align="right"><strong>快递单号：</strong></td>
          <td height="30" align="left">&nbsp;<input name="expressno" type="text" id="expressno" value="" /></td>
        </tr>
        <tr>
          <td align="right">&nbsp;</td>
          <td height="50" align="left">&nbsp;<div class="txtButton">
		  <input type="image" name="Sumitbtn" id="btnSnap" src="../images/jieducm_ConfirmButton.gif" />		 
		  </div></td>
        </tr>			
		<tr>
			<td height="40" colspan="3" align="center" class="orange">注：由于查询会消耗一定系统资源，所以每次查询需间隔10秒。</td>
		</tr>
      </table>
      

    <div class="t_list" id="t_list" style="display:none;">
        <span class="title">&nbsp; 单号生成与查询结果：</span>
        <div id="retData"></div>
    </div>
	
       <div class="Tips">
	  <p><span class="orange rough">随机生成：</span>	  本方式由系统根据所选快递公司单号规则随机生成5条单号。<br />
	  <span class="orange rough">自定义生成：</span><br />
	    1.本方式适合追求近期真实运单号的用户使用；<br />
	    2.您可以将手上已有的真实快递单号完整输入，系统会自动识别并生成邻近的运单号；<br />
	    3.你也可以输入真实运单号（去掉最后的1-3位数字）的前几位数字，系统会自动识别并生成邻近的运单号（单号需要与快递公司匹配）。<br />
	    <span class="orange rough">建议：您可以根据生成运单的收货地址要求接手方淘宝收货地址填写您生成的运单收货信息；淘宝的运单查询只会匹配到达城市，请放心使用！ </span><br />
	    <span class="red rough">注意：</span>为防止恶意，普通会员每天可以使用<span class="orange rough"><%=stup_RndQueryNum%></span>次随机生成和<span class="orange rough"><%=stup_DefQueryNum%></span>次自定义生成，VIP会员均无限！</p>
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