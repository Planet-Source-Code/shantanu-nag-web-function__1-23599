<%@ Language=VBScript %>
<% Option Explicit %>

<%
'*********************************************************************
' ASP Module:        WebFunction Server
' Author:            Shantanu Nag
' Description:       This ASP Web Function will connect to the ACCESS
'						 database and will extract all the Stock Symbols
'						for the selected users.
' Revision History:
' Version       Date        Person              Description
' =======       =====       ================    =================
' 1.0.0         05/2001     Shantanu Nag        Created the ASP page
'*********************************************************************
%>



<%
Dim SQL, rs, cmd
Dim UserName

	'Get the UserName which was passed by the WF Client
	UserName = Request("UserName")
		
	'Formulate the SQL to get the stock symbols for the user
	SQL = "SELECT Stock.StockSymbol " & _
			"FROM [User] INNER JOIN Stock ON User.userID = Stock.userID " & _
			"WHERE UserName = '" & UserName & "'"
			
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open SQL, Application("WFDemo_ConnectionString")
	
	'Write down all the Stock Symbol seperated by a comma
	'The WFClient will parse it and load into an array, from 
	'where it will be loaded to a combo box.
	Do while not rs.eof
		Response.write (rs("StockSymbol"))		
		rs.movenext
		if not rs.eof then
			Response.write(",")
		end if
	Loop

	rs.close
	Set rs = nothing	
%>