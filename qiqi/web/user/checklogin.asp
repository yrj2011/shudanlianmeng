<%  	Sql = "select * from "&jieducm&"_register where username='"&session("useridname")&"'"
				Set rs = Server.CreateObject("Adodb.RecordSet")
				rs.Open Sql,conn,1,1
				IF rs.Eof Then
				  Response.Write("<script>alert('您没有登录，或登录超时，请登录后操作!');window.location.href='../login.asp';</script>")
			     response.End()
				else
				uid=rs("id")
				nowe=rs("snow")
				ismessage=rs("ismessage")
				isdun=rs("isdun")
				isgives=rs("isgives")
				isgive=rs("isgive")
				fabudian=rs("fabudian")
				cunkuan=rs("cunkuan")
				jifei=rs("jifei")
				tjr=Replace(Trim(rs("tjr")),"'","''")
				czm=Replace(Trim(rs("czm")),"'","''")
				level1=Replace(Trim(rs("level1")),"'","''")
				taobao=Replace(Trim(rs("taobao")),"'","''")
				paipai=Replace(Trim(rs("paipai")),"'","''")
				youa=Replace(Trim(rs("youa")),"'","''")
				vip=Replace(Trim(rs("vip")),"'","''")	
				vipdel=rs("vipdel")
				vipdate=rs("vipdate")
			    dst=Replace(Trim(rs("dst")),"'","''")
				dst1=Replace(Trim(rs("dst1")),"'","''")
				codenum=Replace(Trim(rs("codenum")),"'","''")
				tribeid=rs("tribeid")
				login_ip=rs("login_ip")
				rs.close
				end if
		
		if level1="0" then
			 Response.Write("<script>alert('对不起此账号还没有通过管理员审核!');window.location.href='../index.asp';</script>")
			 response.End()
		elseif login_ip<>Request.ServerVariables("REMOTE_ADDR") and login_ip<>"0" then
		     session("useridname")=""
             session("czm")=""
			 Response.Write("<script>alert('你设置的ip和登陆的ip不同!');window.location.href='../index.asp';</script>")
			 response.End()
		end if
								
			sql="select count(*) as weidu1 from "&jieducm&"_china_message where uid='all' and hits=0"
            Set Rs = Server.CreateObject("Adodb.RecordSet")
			Rs.Open Sql,conn,1,1
			if not rs.eof then
			weidu1=rs("weidu1")
			rs.close
			end if
			
			sqlw="select count(*) as weidu from "&jieducm&"_china_message where uid='"&session("useridname")&"' and hits=0"
            Set rs = Server.CreateObject("Adodb.RecordSet")
			rs.Open Sqlw,conn,1,1
			if not rs.eof then
			weidu=rs("weidu")
			rs.close
			end if	
			
		    sqls="select   count(*) as shengsu  from "&jieducm&"_toushu where name='"&session("useridname")&"' "
			Set rs = Server.CreateObject("Adodb.RecordSet")
			rs.Open Sqls,conn,1,1
			if not rs.eof then
			shengsu=rs("shengsu")
			rs.close
			end if
			set rs=nothing
			 %>
			 