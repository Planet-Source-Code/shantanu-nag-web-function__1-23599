<%@ Language=VBScript %>
<% Option Explicit %>
<%
'*********************************************************************
' ASP Module:        WebFunction Server
' Author:            Shantanu Nag
' Description:       This ASP Web Function will connect to the ACCESS
'						 database will validate the username and password
'						combination and will write out TRUE or FALSE.
' Revision History:
' Version       Date        Person              Description
' =======       =====       ================    =================
' 1.0.0         05/2001     Shantanu Nag        Created the ASP page
'*********************************************************************
%>

<%
Dim SQL, rs, cmd
Dim UserName, Passwd

	UserName = Request("UserName")
	Passwd = Request("Passwd")

	SQL = "SELECT UserID FROM User " & _
			"WHERE UserName = '" & UserName & "' " & _
			"AND Password = '" & Passwd & "'"
		
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	rs.open SQL, Application("WFDemo_ConnectionString")
	if not rs.eof then
		Response.Write ("TRUE")
	else
		Response.Write("FALSE")
	End if
	
	rs.close
	Set rs = nothing	

%>