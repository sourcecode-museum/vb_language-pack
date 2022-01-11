VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Language Pack Generator"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox picIcon 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   16
      X2              =   352
      Y1              =   175
      Y2              =   175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   17
      X2              =   351
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Copyright © 2002 Fredisoft Corp."
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   2970
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Contact: indiofu@bol.com.br"
      Height          =   255
      Index           =   4
      Left            =   255
      TabIndex        =   5
      Top             =   2160
      Width           =   5040
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   $"frmAbout.frx":000C
      Height          =   855
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version x.xx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Language Pack Generator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Language Pack Generator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1095
      TabIndex        =   1
      Top             =   135
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  picIcon.Picture = frmMain.Icon
  lblversion = "Version " & App.Major & "." & App.Minor & App.Revision
End Sub
