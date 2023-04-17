<%
'code here
Dim connDB
set connDB = Server.CreateObject("ADODB.Connection")
Dim strConnection
strConnection = "Provider=SQLOLEDB.1;Data Source=TUNZTUNZ\SQLEXPRESS;Database=QLNV;User Id=sa;Password=123456"
connDB.ConnectionString = strConnection
%>