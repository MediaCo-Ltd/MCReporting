<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>
 
<%
Session("LoggedOff") = True
    
Session.Abandon
Response.Redirect("login.asp")
%>
