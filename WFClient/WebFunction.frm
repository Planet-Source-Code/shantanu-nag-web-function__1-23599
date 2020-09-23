VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form WebFunction 
   Caption         =   "Form1"
   ClientHeight    =   645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet 
      Left            =   120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "WebFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************
' VB Module:         WebFunction
' Author:            Shantanu Nag
' Description:       This module hosts the Internet Control and uses it for
'                    Communicating with the WF Server.
' Revision History:
' Version       Date        Person              Description
' =======       =====       ================    =================
' 1.0.0         05/2001     Shantanu Nag        Created the module
'*********************************************************************

'A Constant which points to the Web Folder for the Web Functions
'This URL should work for the samples.
Const WFLocation = "http://www.xsolvenet.com/Articles/WebFunction/WFServer/"
Public WFName As String
Public DisplayStatus As String


'********************************************************************
'Purpose:   Executes the Web Function based on the URL. The URL is
'           linked to the ASP page where the function is implemented.
'           After the function is executed on the Web Server the result
'           then returned back to the calling Web Function.
'Inputs:    The URL to the Web Function on the Web Server
'Returns:   String Value
'DependsOn: The ASP Page which implements the Web function
'Effects:   None
'********************************************************************
' Revision History:
' Date          Person              Description
' ============  ================    ======================
' 07/05/2001    Shantanu Nag        Created the Method.
'********************************************************************

Public Function RunWebFunction(ByVal URL As String) As String
    Inet.URL = WFLocation & URL
    RunWebFunction = Inet.OpenURL
End Function


'********************************************************************
'Purpose:   Executes the Hello World Web Function. The URL to the
'           implementation file is WF_HelloWorld.asp
'Inputs:    None
'Returns:   String Value
'DependsOn: The ASP Page which implements the Hello World Web function
'Effects:   None
'********************************************************************
' Revision History:
' Date          Person              Description
' ============  ================    ======================
' 07/05/2001    Shantanu Nag        Created the Method.
'********************************************************************
Public Function HelloWorld() As String
                                
    WFName = "WF_HelloWorld.asp"
    HelloWorld = RunWebFunction(WFName)

End Function

'********************************************************************
'Purpose:   Executes the Validate User Web Function. The URL to the
'           implementation file is WF_ValidateUser.asp
'Inputs:    UserName and Password from the Login Form
'Returns:   Boolean
'DependsOn: The ASP Page which implements the Validate User Function
'Effects:   None
'********************************************************************
' Revision History:
' Date          Person              Description
' ============  ================    ======================
' 07/05/2001    Shantanu Nag        Created the Method.
'********************************************************************
Public Function ValidateUser(ByVal UserName As String, _
                            ByVal Passwd As String) As Boolean
                                
    WFName = "WF_ValidateUser.asp?UserName=" & UserName & "&Passwd=" & Passwd
    ValidateUser = CBool(RunWebFunction(WFName))

End Function


'********************************************************************
'Purpose:   Executes the Get Stock Symbol Web Function. The URL to the
'           implementation file is WF_GetStockSymbol.asp
'Inputs:    UserName from the Login Form
'Returns:   String array
'DependsOn: The ASP Page which implements the Get Stock Symbol Web Function
'Effects:   None
'********************************************************************
' Revision History:
' Date          Person              Description
' ============  ================    ======================
' 07/05/2001    Shantanu Nag        Created the Method.
'********************************************************************
Public Function GetStockSymbol(ByVal UserName As String) As String()
Dim temp As String

    WFName = "WF_GetStockSymbol.asp?UserName=" & UserName
    temp = RunWebFunction(WFName)
    GetStockSymbol = Split(temp, ",")
End Function




'********************************************************************
'Purpose:   This Inet Event is used for displaying the messages
'Inputs:    State
'Returns:   None
'DependsOn: None
'Effects:   DisplayStatus variable
'********************************************************************
' Revision History:
' Date          Person              Description
' ============  ================    ======================
' 07/05/2001    Shantanu Nag        Created the Method.
'********************************************************************
Private Sub Inet_StateChanged(ByVal State As Integer)

    Select Case State
        Case icConnected
            DisplayStatus = "Connected"

        Case icConnecting
            DisplayStatus = "Connecting"

        Case icDisconnected
            DisplayStatus = "Disconnected"

        Case icDisconnecting
            DisplayStatus = "Disconnecting"

        Case icError
            DisplayStatus = "Error: " & Inet.ResponseInfo

        Case icReceivingResponse
            DisplayStatus = "Receiving response"

        Case icRequesting
            DisplayStatus = "Requesting"

        Case icRequestSent
            DisplayStatus = "Request Sent"

        Case icResolvingHost
            DisplayStatus = "Resolving host"

        Case icResponseCompleted
            DisplayStatus = "Response completed"

        Case icResponseReceived
            DisplayStatus = "Response received"
    End Select
    
    frmSample.lblMsg.Caption = DisplayStatus
    
End Sub


