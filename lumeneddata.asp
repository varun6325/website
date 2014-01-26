<%@Language=VBScript%>
<!--#include file="dbconnect.asp"-->


<%



'********************************  TO SAVE DATA IN DATABASE TABLE *****************
   set objrs=server.createobject("ADODB.recordset")   
   objrs.open "lumened", Objconn,2,3     

objrs("email") = Request.Form("email")


objrs.Update
objrs.Close
Set objrs = Nothing
Response.Write("THANK YOU FOR YOUR RESPONSE.. <br>   ")
Response.Write(".....Please Wait.....")
Response.AddHeader "REFRESH","5;URL=lumened.asp"
%>


