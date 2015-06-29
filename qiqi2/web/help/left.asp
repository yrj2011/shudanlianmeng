<DIV class=menu>
      <UL>
        <LI><A 
        href="help10.asp?action=82"><STRONG>初识刷客</STRONG></A> 
        <UL>
          <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (82) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
					
			  %>
   
          <li> <a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a>
         
        </LI> <%
			  	Rs.MoveNext
				Loop
				End IF
			  %></UL>
        <LI><A class=hide 
        href="help10.asp?action=84"><STRONG>刷客类型</STRONG></A> 
        <UL>
         
          <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (84) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
					
			  %>
   
          <li> <a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a>
         
        </LI> <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
         </UL>
        <LI><A class=hide 
        href="help10.asp?action=86"><STRONG>刷客相关</STRONG></A> 
        <UL>
          
           <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (86) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
					
			  %>
   
          <li> <a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a>
         
        </LI> <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
          </UL>
        <LI><A class=hide 
        href="help10.asp?action=83"><STRONG>互刷帮助</STRONG></A> 
        <UL>
           <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (83) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
					
			  %>
   
          <li> <a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a>
         
        </LI> <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
          </UL></LI>
           <LI><A class=hide 
        href="help10.asp?action=85"><STRONG>刷客宝典</STRONG></A> 
        <UL>
           <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (85) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
					
			  %>
   
          <li> <a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a>
         
        </LI> <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
          </UL></LI>
          <LI><A class=hide 
        href="help10.asp?action=87"><STRONG>高手必读</STRONG></A> 
        <UL>
           <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (87) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
					
			  %>
   
          <li> <a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a>
         
        </LI> <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
          </UL></LI>
  </UL>
</DIV>