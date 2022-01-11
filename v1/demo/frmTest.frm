VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Just a test"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "This is another test, just to be sure that it works with multiple forms."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Example: Use of the Language Pack Class Module
' // Developed by Frederico Machado (indiofu@bol.com.br)
' // Vote for me if you like it please!
' //
' // I don't know if it have bugs, cause I haven't tested
' // it deeply. If you find any bug in the Packer or in
' // the Class Module, please, let me know about it.
' // Thank you!
' // P.S.: Please, don't forget to give me some credit
' // if you use this code in your own VB softwares.
' /////////////////////////////////////////////////////////

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  ' If there is a current language pack loaded
  ' then set it in this form too.
  If cLanguage.sCurrentFile <> "" Then
    cLanguage.SetLanguageInForm Me
  End If
End Sub
