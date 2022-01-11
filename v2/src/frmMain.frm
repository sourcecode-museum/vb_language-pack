VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Language Pack Generator"
   ClientHeight    =   4950
   ClientLeft      =   1605
   ClientTop       =   3615
   ClientWidth     =   9585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Pack"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tvObjects 
      Height          =   2055
      Left            =   3960
      TabIndex        =   26
      Top             =   1320
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3625
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvForms 
      Height          =   2055
      Left            =   120
      TabIndex        =   25
      Top             =   1320
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3625
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4680
      Top             =   -360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1372
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   6480
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Pack"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtNewToolTip 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   4440
      Width           =   6495
   End
   Begin VB.TextBox txtNewCaption 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   3960
      Width           =   6495
   End
   Begin VB.TextBox txtToolTip 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txtCaption 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton cmdBrowseOut 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtOutputFile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   5535
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load VBP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtProject 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2005 Frederico Machado."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   8040
      TabIndex        =   24
      Top             =   3120
      Width           =   1515
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Language Pack Generator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1095
      Left            =   8145
      TabIndex        =   23
      Top             =   1785
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Language Pack Generator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   8160
      TabIndex        =   22
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003 Frederico Machado."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8055
      TabIndex        =   21
      Top             =   3135
      Width           =   1515
   End
   Begin VB.Label Label8 
      Caption         =   "New ToolTip:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "New Caption:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4005
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "ToolTip:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   3525
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Caption:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3525
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Objects/Controls:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Forms:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Output File:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   645
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   527
      X2              =   527
      Y1              =   16
      Y2              =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   528
      X2              =   528
      Y1              =   17
      Y2              =   314
   End
   Begin VB.Label Label1 
      Caption         =   "VB Project:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   285
      Width           =   975
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuGrid 
         Caption         =   "&Grid View"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // VB Language Pack Generator 1.30
' // Developed by Frederico Machado (fredisoft@terra.com.br)
' // Vote for me if you like it please!
' ////////////////////////////////////////////////////////

Option Explicit

Dim sVer As String * 4
Dim sProj As String * 32
Dim iForms As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAbout_Click()
  FormSplash.Show 1
End Sub

Private Sub cmdBrowse_Click()
  
  cmDialog.FileName = ""
  cmDialog.DialogTitle = "Open a VB Project"
  cmDialog.Filter = "VB Project (*.vbp)|*.vbp|"
  cmDialog.Flags = &H4 Or &H1000
  cmDialog.ShowOpen
  sVBProject = cmDialog.FileName
  If sVBProject = "" Then Exit Sub
  sPrjFolder = Left(sVBProject, InStrRev(sVBProject, "\"))
  txtProject = sVBProject

End Sub

Private Sub cmdBrowseOut_Click()

  cmDialog.FileName = ""
  cmDialog.DialogTitle = "Save Language Pack"
  cmDialog.Filter = "Language Pack (*.lpk)|*.lpk|"
  cmDialog.Flags = &H4 Or &H1000
  cmDialog.ShowSave
  sOutputFile = cmDialog.FileName
  If sOutputFile = "" Then Exit Sub
  txtOutputFile = sOutputFile

End Sub

Private Sub SavePack(sFile As String)
    On Local Error GoTo ERRO
    
    If sOutputFile = "" Then
        MsgBox "You need to specify an Output file to the Pack", vbCritical
        txtOutputFile.SetFocus
        Exit Sub
    End If

    iForms = UBound(FormProp)
    
    Dim i As Integer, j As Integer
    
    sProj = Encrypt(sProj)
    
    For i = 1 To iForms
        FormProp(i).FileName = Encrypt(FormProp(i).FileName)
        FormProp(i).Name = Encrypt(FormProp(i).Name)
        FormProp(i).Caption = Encrypt(FormProp(i).Caption)
        For j = 1 To FormProp(i).ObjectNumber
            FormProp(i).objProp(j).Name = Encrypt(FormProp(i).objProp(j).Name)
            FormProp(i).objProp(j).Caption = Encrypt(FormProp(i).objProp(j).Caption)
            FormProp(i).objProp(j).NewCaption = Encrypt(FormProp(i).objProp(j).NewCaption)
            FormProp(i).objProp(j).ToolTip = Encrypt(FormProp(i).objProp(j).ToolTip)
            FormProp(i).objProp(j).NewToolTip = Encrypt(FormProp(i).objProp(j).NewToolTip)
        Next
    Next
    
    For i = 1 To iStrings
        LPGStrings(i).Name = Encrypt(LPGStrings(i).Name)
        LPGStrings(i).String = Encrypt(LPGStrings(i).String)
    Next
    
    Open sFile For Binary Access Write Lock Write As #1
        Put #1, 1, sVer
        Put #1, , sProj
        Put #1, , iForms
        Put #1, , FormProp
        Put #1, , iStrings
        Put #1, , LPGStrings
    Close #1
    
    sProj = Decrypt(sProj)
    
    For i = 1 To iForms
        FormProp(i).FileName = Decrypt(FormProp(i).FileName)
        FormProp(i).Name = Decrypt(FormProp(i).Name)
        FormProp(i).Caption = Decrypt(FormProp(i).Caption)
        For j = 1 To FormProp(i).ObjectNumber
            FormProp(i).objProp(j).Name = Decrypt(FormProp(i).objProp(j).Name)
            FormProp(i).objProp(j).Caption = Decrypt(FormProp(i).objProp(j).Caption)
            FormProp(i).objProp(j).NewCaption = Decrypt(FormProp(i).objProp(j).NewCaption)
            FormProp(i).objProp(j).ToolTip = Decrypt(FormProp(i).objProp(j).ToolTip)
            FormProp(i).objProp(j).NewToolTip = Decrypt(FormProp(i).objProp(j).NewToolTip)
        Next
    Next
    
    For i = 1 To iStrings
        LPGStrings(i).Name = Decrypt(LPGStrings(i).Name)
        LPGStrings(i).String = Decrypt(LPGStrings(i).String)
    Next
    
    MsgBox "Language Pack successfuly created.", vbInformation
    
    Exit Sub
    
ERRO:
    MsgBox "There was an error while creating the Language Pack.", vbCritical, "Error"
    Kill sOutputFile
End Sub

Private Sub cmdGenerate_Click()
  SavePack sOutputFile
End Sub

Private Sub cmdHelp_Click()
    On Local Error GoTo ERRO
    ShellExecute 0, "open", App.Path & "\help\index.htm", "", "", 0
    Exit Sub
ERRO:
    MsgBox "The help file was not found in the help folder.", vbCritical
End Sub

Private Sub LoadPack()
    Dim i As Integer, j As Integer, sFile As String, sProj2 As String
    
    cmDialog.FileName = ""
    cmDialog.DialogTitle = "Open Language Pack"
    cmDialog.Filter = "Language Pack (*.lpk)|*.lpk|"
    cmDialog.Flags = &H4 Or &H1000
    cmDialog.ShowOpen
    sFile = cmDialog.FileName
    If sFile = "" Then Exit Sub
    
    Open sFile For Binary Access Read Lock Write As #1
        Get #1, 1, sVer
        Get #1, , sProj
        Get #1, , iForms
        ReDim FormProp(iForms)
        Get #1, , FormProp
        Get #1, , iStrings
        ReDim LPGStrings(iStrings)
        Get #1, , LPGStrings
    Close #1
    
    sProj = Decrypt(sProj)
    
    For i = 1 To iForms
        FormProp(i).FileName = Decrypt(FormProp(i).FileName)
        FormProp(i).Name = Decrypt(FormProp(i).Name)
        FormProp(i).Caption = Decrypt(FormProp(i).Caption)
        For j = 1 To FormProp(i).ObjectNumber
            FormProp(i).objProp(j).Name = Decrypt(FormProp(i).objProp(j).Name)
            FormProp(i).objProp(j).Caption = Decrypt(FormProp(i).objProp(j).Caption)
            FormProp(i).objProp(j).NewCaption = Decrypt(FormProp(i).objProp(j).NewCaption)
            FormProp(i).objProp(j).ToolTip = Decrypt(FormProp(i).objProp(j).ToolTip)
            FormProp(i).objProp(j).NewToolTip = Decrypt(FormProp(i).objProp(j).NewToolTip)
        Next
    Next
    
    For i = 1 To iStrings
        LPGStrings(i).Name = Decrypt(LPGStrings(i).Name)
        LPGStrings(i).String = Decrypt(LPGStrings(i).String)
    Next
    
    tvForms.Nodes.Clear
    tvObjects.Nodes.Clear
    
    sProj2 = Trim(sProj)
    tvForms.Nodes.Add , , "Project", sProj2, 1
    tvForms.Nodes.Add "Project", tvwChild, "Forms", "Forms", 2
    
    For i = 1 To iForms
        tvForms.Nodes.Add "Forms", tvwChild, FormProp(i).FileName, FormProp(i).FileName, 4
    Next
    
    tvForms.Nodes.Add "Project", tvwChild, "nStrings", "Strings", 2
    tvForms.Nodes.Add "nStrings", tvwChild, "Strings", "Strings"
    
    tvForms.Nodes.Item(1).Expanded = True
    tvForms.Nodes.Item(2).Expanded = True
    tvForms.Nodes.Item(tvForms.Nodes.Count - 1).Expanded = True
    
    bPackLd = True
End Sub

Private Sub cmdLoad_Click()
  
  Dim sLine As String, iCount As Integer, sProj2 As String
  
  If sVBProject = "" Then
    MsgBox "You need to select a VB Project to load forms and controls.", vbCritical
    txtProject.SetFocus
    Exit Sub
  End If
  
  Caption = "VB Language Pack Generator - Loading ..."

  Me.MousePointer = 11
  If Not bPackLd Then
    tvForms.Nodes.Clear
    tvObjects.Nodes.Clear
  End If

  Dim sForms As String, sFArray() As String

  ' Scan the project for forms
  Open sVBProject For Input As #1
    Do
      Input #1, sLine
      If LCase(Left(sLine, 4)) = "name" Then
        sProj = Replace(Mid$(sLine, 6), Chr$(34), "")
        sProj2 = Trim(sProj)
        If Not bPackLd Then tvForms.Nodes.Add , , "Project", sProj2, 1
      ElseIf LCase(Left(sLine, 4)) = "form" Then
        sForms = sForms & Mid$(sLine, 6) & "|"
        iCount = iCount + 1
      End If
    Loop Until EOF(1)
  Close #1
  
  sFArray = Split(sForms, "|")
  If Not bPackLd Then tvForms.Nodes.Add "Project", tvwChild, "Forms", "Forms", 2
    
  ReDim Preserve FormProp(iCount)
  
  Dim i As Integer
  ' Get the properties and objects of each form
  For i = 1 To iCount
    If bPackLd = False Or bFormExists(sFArray(i - 1)) = False Then
        tvForms.Nodes.Add "Forms", tvwChild, sFArray(i - 1), sFArray(i - 1), 4
        FormProp(i).FileName = sFArray(i - 1)
    End If
    ReadFormFile sPrjFolder & sFArray(i - 1), i
  Next
  
  If Not bPackLd Then
    tvForms.Nodes.Add "Project", tvwChild, "nStrings", "Strings", 2
    tvForms.Nodes.Add "nStrings", tvwChild, "Strings", "Strings"
  End If
  
  tvForms.Nodes.Item(1).Expanded = True
  tvForms.Nodes.Item(2).Expanded = True
  tvForms.Nodes.Item(tvForms.Nodes.Count - 1).Expanded = True
  
  Me.MousePointer = 0
  
  Caption = "VB Language Pack Generator"

End Sub

Sub CleanTexts()
  txtCaption = "": txtToolTip = "": txtNewCaption = "": txtNewToolTip = ""
End Sub

Private Sub cmdOpen_Click()
    LoadPack
End Sub

Private Sub Form_Load()
  txtCaption.Enabled = False
  txtNewCaption.Enabled = False
  txtToolTip.Enabled = False
  txtNewToolTip.Enabled = False
  txtCaption.BackColor = &H8000000F
  txtNewCaption.BackColor = &H8000000F
  txtToolTip.BackColor = &H8000000F
  txtNewToolTip.BackColor = &H8000000F
  sVer = Left(App.Major & "." & App.Minor & App.Revision, 4)
End Sub

Private Sub mnuGrid_Click()
    If bStrings Then
        Call frmGrid.Carregar(tgStrings)
    ElseIf iCurrForm > 0 Then
        Call frmGrid.Carregar(tgForms)
    End If
End Sub

Private Sub tvForms_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Text = "Forms" Then
    Node.Image = 2
  End If
End Sub

Private Sub tvForms_Expand(ByVal Node As MSComctlLib.Node)
  If Node.Text = "Forms" Then
    Node.Image = 3
  End If
End Sub

Private Sub tvForms_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And bStrings = True Then
        PopupMenu mnuPopup
        Exit Sub
    End If
    If Button = 2 Then
        If iCurrForm = 0 Then Exit Sub
        PopupMenu mnuPopup
    End If
End Sub

Private Sub tvForms_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim ControlName As String
  Dim i As Integer, bAdded(6) As Boolean
  
  iCurrForm = 0
  bStrings = False
  
  If Node.Index < 3 Then Exit Sub
  If Node.Key = "nStrings" Then Exit Sub
  tvObjects.Nodes.Clear
  CleanTexts
  
  If Node.Key = "Strings" Then
    bStrings = True
    tvObjects.Nodes.Add , , "AddString", "Add String"
    If iStrings > 0 Then
        tvObjects.Nodes.Add , , "Strings", "Strings"
        For i = 1 To iStrings
            tvObjects.Nodes.Add "Strings", tvwChild, LPGStrings(i).Name, LPGStrings(i).Name
            If i = 1 Then tvObjects.Nodes.Item(tvObjects.Nodes.Count - 1).Expanded = True
        Next
    End If
    GoTo Jump
  End If
  
  iCurrForm = FormIndex(Node.Text)
  tvObjects.Nodes.Add , , "Form", FormProp(iCurrForm).Name
  
  ' List the objects and its properties in the list
  For i = 1 To FormProp(iCurrForm).ObjectNumber
    ControlName = FormProp(iCurrForm).objProp(i).Name
    If LCase(FormProp(iCurrForm).objProp(i).Type) = "checkbox" Then
      If Not bAdded(0) Then
        tvObjects.Nodes.Add , , "CheckBox", "CheckBox"
        bAdded(0) = True
      End If
      tvObjects.Nodes.Add "CheckBox", tvwChild, FormProp(iCurrForm).objProp(i).Name, FormProp(iCurrForm).objProp(i).Name
    ElseIf LCase(FormProp(iCurrForm).objProp(i).Type) = "commandbutton" Then
      If Not bAdded(1) Then
        tvObjects.Nodes.Add , , "CommandButton", "CommandButton"
        bAdded(1) = True
      End If
      tvObjects.Nodes.Add "CommandButton", tvwChild, FormProp(iCurrForm).objProp(i).Name, FormProp(iCurrForm).objProp(i).Name
    ElseIf LCase(FormProp(iCurrForm).objProp(i).Type) = "frame" Then
      If Not bAdded(2) Then
        tvObjects.Nodes.Add , , "Frame", "Frame"
        bAdded(2) = True
      End If
      tvObjects.Nodes.Add "Frame", tvwChild, FormProp(iCurrForm).objProp(i).Name, FormProp(iCurrForm).objProp(i).Name
    ElseIf LCase(FormProp(iCurrForm).objProp(i).Type) = "label" Then
      If Not bAdded(3) Then
        tvObjects.Nodes.Add , , "Label", "Label"
        bAdded(3) = True
      End If
      tvObjects.Nodes.Add "Label", tvwChild, FormProp(iCurrForm).objProp(i).Name & "_" & FormProp(iCurrForm).objProp(i).Index, FormProp(iCurrForm).objProp(i).Name
    ElseIf LCase(FormProp(iCurrForm).objProp(i).Type) = "optionbutton" Then
      If Not bAdded(4) Then
        tvObjects.Nodes.Add , , "OptionButton", "OptionButton"
        bAdded(4) = True
      End If
      tvObjects.Nodes.Add "OptionButton", tvwChild, FormProp(iCurrForm).objProp(i).Name, FormProp(iCurrForm).objProp(i).Name
    ElseIf LCase(FormProp(iCurrForm).objProp(i).Type) = "menu" Then
      If Not bAdded(5) Then
        tvObjects.Nodes.Add , , "Menu", "Menu"
        bAdded(5) = True
      End If
      tvObjects.Nodes.Add "Menu", tvwChild, FormProp(iCurrForm).objProp(i).Name, FormProp(iCurrForm).objProp(i).Name
    ElseIf LCase(FormProp(iCurrForm).objProp(i).Type) = "sstab" Then
      If Not bAdded(6) Then
        tvObjects.Nodes.Add , , "SSTab", "SSTab"
        bAdded(6) = True
      End If
      tvObjects.Nodes.Add "SSTab", tvwChild, ControlName, ControlName
    End If
  Next
  
Jump:
  
  txtCaption.Enabled = False
  txtNewCaption.Enabled = False
  txtToolTip.Enabled = False
  txtNewToolTip.Enabled = False
  txtCaption.BackColor = &H8000000F
  txtNewCaption.BackColor = &H8000000F
  txtToolTip.BackColor = &H8000000F
  txtNewToolTip.BackColor = &H8000000F
End Sub

Private Sub tvObjects_NodeClick(ByVal Node As MSComctlLib.Node)
  CleanTexts
  
  Dim sTName As String, sTString As String
  
  If bStrings Then
    If Node.Text = "Add String" Then
      sTName = InputBox("Enter the name of the string:", "Add String")
      sTString = InputBox("Enter the string:", "String")
      If sTName <> "" And sTString <> "" Then
          iStrings = iStrings + 1
          ReDim Preserve LPGStrings(iStrings)
          LPGStrings(iStrings).Name = sTName
          LPGStrings(iStrings).String = sTString
          tvForms_NodeClick tvForms.Nodes(tvForms.Nodes.Count)
      End If
    ElseIf Node.Key = "Strings" Then
        Exit Sub
    Else
      sTString = InputBox("Enter the string:", "String", GetString(Node.Text))
      If sTString <> "" Then
        LPGStrings(GetStringId(Node.Text)).String = sTString
      End If
    End If
    
    Exit Sub
  End If
  
  iCurrObj = Node.Index - 1
  
  txtCaption.Enabled = True
  txtNewCaption.Enabled = True
  txtCaption.BackColor = vbWhite
  txtNewCaption.BackColor = vbWhite
  
  ' It is the form, get its caption and new caption
  ' Forms don't have tooltip, then disable it
  If iCurrObj = 0 Then
    txtCaption = FormProp(iCurrForm).Caption
    txtNewCaption = FormProp(iCurrForm).NewCaption
    txtToolTip.Enabled = False
    txtNewToolTip.Enabled = False
    txtToolTip.BackColor = &H8000000F
    txtNewToolTip.BackColor = &H8000000F
    Exit Sub
  End If

  Dim i As Integer
  If Node.Children > 0 Then
    txtCaption.Enabled = False
    txtNewCaption.Enabled = False
    txtToolTip.Enabled = False
    txtNewToolTip.Enabled = False
    txtCaption.BackColor = &H8000000F
    txtNewCaption.BackColor = &H8000000F
    txtToolTip.BackColor = &H8000000F
    txtNewToolTip.BackColor = &H8000000F
    Exit Sub
  End If
  
  ' It is a menu, get its caption and new caption
  ' Menus don't have tooltip too.
  If Node.Parent.Key = "Menu" Then
    For i = 0 To FormProp(iCurrForm).ObjectNumber
      If FormProp(iCurrForm).objProp(i).Name = Node.Text Then
        txtCaption = FormProp(iCurrForm).objProp(i).Caption
        txtNewCaption = FormProp(iCurrForm).objProp(i).NewCaption
        If txtCaption = "-" Then
          txtCaption = "(separator)"
          txtNewCaption = "(separator)"
          txtCaption.Enabled = False
          txtNewCaption.Enabled = False
          txtCaption.BackColor = &H8000000F
          txtNewCaption.BackColor = &H8000000F
        End If
        txtToolTip.Enabled = False
        txtNewToolTip.Enabled = False
        txtToolTip.BackColor = &H8000000F
        txtNewToolTip.BackColor = &H8000000F
        iCurrObj = i
        Exit For
      End If
    Next
    Exit Sub
  End If
  
  ' Get caption, new caption, tooltip and new tooltip
  For i = 0 To FormProp(iCurrForm).ObjectNumber
    If FormProp(iCurrForm).objProp(i).Name = Node.Text Then
      txtToolTip.Enabled = True
      txtNewToolTip.Enabled = True
      txtToolTip.BackColor = vbWhite
      txtNewToolTip.BackColor = vbWhite
      txtCaption = FormProp(iCurrForm).objProp(i).Caption
      txtToolTip = FormProp(iCurrForm).objProp(i).ToolTip
      txtNewCaption = FormProp(iCurrForm).objProp(i).NewCaption
      txtNewToolTip = FormProp(iCurrForm).objProp(i).NewToolTip
      iCurrObj = i
      Exit For
    End If
  Next
End Sub

Private Sub txtNewCaption_LostFocus()
  ' Set the new caption
  If iCurrObj = 0 Then
    ' It is the form caption
    FormProp(iCurrForm).NewCaption = txtNewCaption
    Exit Sub
  End If
  FormProp(iCurrForm).objProp(iCurrObj).NewCaption = txtNewCaption
End Sub

Private Sub txtNewToolTip_LostFocus()
  ' Set the new tooltip
  FormProp(iCurrForm).objProp(iCurrObj).NewToolTip = txtNewToolTip
End Sub

Private Sub txtOutputFile_LostFocus()
  ' Set the output file
  sOutputFile = txtOutputFile
End Sub

Private Sub txtProject_LostFocus()
  ' Set the project file
  sVBProject = txtProject
End Sub
