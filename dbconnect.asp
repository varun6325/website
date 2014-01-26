<%
'*************************   Create a connection to the database  **************************** 
'dim Objconn,filepath
'filepath=server.mapPath("tis.mdb")
Set Objconn=server.createobject("ADODB.Connection")
'Objconn.open "Provider=SQLNCLI;Server=192.168.1.3;Database=tend1;pooling=yes;min pool size=1;max pool size=50;connection reset=true;uid=sa;pwd=tdi@123;Trusted_Connection=yes"
'Objconn.open "DRIVER={SQL Server};Server=192.168.1.29; database=tend1; UID=sa; PWD=admin@123"

Objconn.open "DSN=con;UID=;PWD=;DATABASE=tis;APP=ASP script;"

'**************************************  End  ***********************************************
%>