<%@ Language=VBScript %>
<% Option Explicit %>
<%
'*********************************************************************
' ASP Module:        WebFunction Server
' Author:            Shantanu Nag
' Description:       This ASP Web Function will write "HELLO WORLD".
'						The WFClient will parse the result and display
'						in a message box
' Revision History:
' Version       Date        Person              Description
' =======       =====       ================    =================
' 1.0.0         05/2001     Shantanu Nag        Created the ASP page
'*********************************************************************
%>


<%
Response.Write("HELLO WORLD")
%>

