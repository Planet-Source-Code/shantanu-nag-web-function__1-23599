VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1950
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   753.624
   ScaleMode       =   0  'User
   ScaleWidth      =   3391.473
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbMachine 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Text            =   "Select Stock"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Text            =   "John"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2340
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      TabIndex        =   3
      Text            =   "john"
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Stock Symbol"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************
' VB Form Module:    Login Form
' Author:            Shantanu Nag
' Description:       This form validates the username and password
'                    using a Web Function and then displays all the
'                    all the Stock Symbols for the validated users by
'                    calling another Web Function.
' Revision History:
' Version       Date        Person              Description
' =======       =====       ================    =================
' 1.0.0         05/2001     Shantanu Nag        Created the module
'*********************************************************************

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

'********************************************************************
'Purpose:   This click event takes the username and password and then
'           calls the web function (ValidateUser), which returns a boolean
'           If the user is validated then another web function (GetStockSymbol)
'           is called which return a string array of Stock Symbol which is loaded
'           to the combo box.
'Returns:   String Value
'DependsOn: The ASP Page which implements the Web function
'Effects:   None
'********************************************************************
' Revision History:
' Date          Person              Description
' ============  ================    ======================
' 07/05/2001    Shantanu Nag        Created the Method.
'********************************************************************

Private Sub cmdOK_Click()
    Dim Passwd As String
    Dim UserName As String
    Dim arryMC() As String
    Dim i As Integer
    Dim msg As String
    
    
    UserName = Trim(txtUserName.Text)
    Passwd = Trim(txtPassword.Text)
    
    'check for correct password
     If WebFunction.ValidateUser(UserName, Passwd) Then
         LoginSucceeded = True
         msg = "User is validated! " & vbCrLf & "Press Ok to download your Stock Symbols"
         MsgBox (msg)
         UserName = txtUserName.Text
         arryMC = WebFunction.GetStockSymbol(UserName)
         cmbMachine.Enabled = True
         i = 0
         cmbMachine.Clear
         cmbMachine.Text = "Select Stock Symbol"
         Do While i <= UBound(arryMC)
             cmbMachine.AddItem arryMC(i)
             i = i + 1
         Loop
     Else
         MsgBox "Invalid User! Please try again.", , "Login"
         txtPassword.SetFocus
         SendKeys "{Home}+{End}"
     End If
        
End Sub


