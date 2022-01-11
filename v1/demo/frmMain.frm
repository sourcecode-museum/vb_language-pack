VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Language pack test form"
   ClientHeight    =   4080
   ClientLeft      =   1785
   ClientTop       =   4170
   ClientWidth     =   8910
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   8910
   Begin VB.CheckBox Check1 
      Caption         =   "Something to check"
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   2400
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      Caption         =   "My Other Option"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   3000
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "My Option"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   5400
      TabIndex        =   8
      Top             =   240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2990
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Test"
      TabPicture(0)   =   "frmMain.frx":08D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "More Test"
      TabPicture(1)   =   "frmMain.frx":08EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Last Tab"
      TabPicture(2)   =   "frmMain.frx":090A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open frmTest"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enumerate Language Packs"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.ListBox lstLangPacks 
      Height          =   840
      ItemData        =   "frmMain.frx":0926
      Left            =   240
      List            =   "frmMain.frx":0928
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Language Pack"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Another Label with index"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   13
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Label with index"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   $"frmMain.frx":092A
      Height          =   675
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "You can make labels or buttons with tooltiptext."
      Height          =   195
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Test TooltipText..."
      Top             =   1800
      Width           =   3330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This is a test."
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   930
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuTest 
         Caption         =   "You can translate menus.."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTest1 
      Caption         =   "&Test"
      Begin VB.Menu mnuIndex 
         Caption         =   "Index 0"
         Index           =   0
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "Index 1"
         Index           =   1
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "Index 2"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
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
  ' Just to be sure if there is Language Packs loaded or if the user selected one
  If lstLangPacks.ListCount = 0 Or lstLangPacks.ListIndex = -1 Then Exit Sub
  
  ' Lets load the entire language pack. It doesn't apply the language pack in the form.
  cLanguage.LoadLanguagePack App.Path & "\packs\" & lstLangPacks.List(lstLangPacks.ListIndex)
  ' Now it applies the language pack in the form
  cLanguage.SetLanguageInForm Me
End Sub

Private Sub Command2_Click()
  ' Clear the listbox. If we clicked in Command2 more than one time, the packs don't repeat in it.
  lstLangPacks.Clear
  
  Dim sTmp As String, sTmpArray() As String, i As Integer
  
  ' Set the temp variable with the function that returns the packs found separated by |
  sTmp = cLanguage.EnumLanguagePacks(App.Path & "\packs", "*.lpk")
  ' Lets split the temp variable into the temp array
  sTmpArray = Split(sTmp, "|")
  ' Lets put the file into the listbox
  For i = 0 To UBound(sTmpArray)
    ' Just to be sure that it's not empty
    If sTmpArray(i) <> "" Then lstLangPacks.AddItem sTmpArray(i)
  Next
End Sub

Private Sub Command3_Click()
  ' Tcharan!
  frmTest.Show
End Sub

Private Sub Command4_Click()
  ' Don't leave, please! :)
  End
End Sub

Private Sub mnuExit_Click()
  ' Don't leave, please! :)
  End
End Sub
