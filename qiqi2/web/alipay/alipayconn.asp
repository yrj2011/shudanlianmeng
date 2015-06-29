<%
jieducm="jieducm_ping"
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.open "driver={SQL Server};server=(local);uid=mingceshi;pwd=mingceshi;database=mingceshi"
'Conn.Open "provider=sqloledb;data source=61.158.138.37,1433;User ID=a0628223712;pwd=23405500;Initial Catalog=a0628223712"
  
set Rsstup=server.createobject("adodb.recordset")
sql="select * from "&jieducm&"_stup"
Rsstup.open sql,conn,1,1
if not (Rsstup.bof) then
stupurl =rsstup("url")
stupzfb=rsstup("zfb")
stupwushi=rsstup("wushi")
stupyibai=rsstup("yibai")
end if
%>