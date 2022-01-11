VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5040
   ClientLeft      =   1605
   ClientTop       =   3615
   ClientWidth     =   10020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Carregar"
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
      Left            =   8310
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Gerar Pacote"
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
      Left            =   8310
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tvObjects 
      Height          =   2055
      Left            =   3960
      TabIndex        =   19
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
      TabIndex        =   18
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
      Caption         =   "Ajuda"
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
      Left            =   8310
      TabIndex        =   9
      Top             =   4500
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
      Left            =   8310
      TabIndex        =   8
      Top             =   4020
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
      Top             =   4545
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
      Top             =   4230
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3795
      Width           =   6495
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3480
      Width           =   6495
   End
   Begin VB.CommandButton cmdBrowseOut 
      Caption         =   "..."
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
      Left            =   6660
      TabIndex        =   3
      Top             =   600
      Width           =   345
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
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
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
      Left            =   6660
      TabIndex        =   1
      Top             =   240
      Width           =   345
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
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gerador de Pacotes Multi-Linguagens"
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
      Height          =   1005
      Index           =   10
      Left            =   7920
      TabIndex        =   23
      Top             =   1335
      Width           =   1995
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004 Insígnia Desenvolvimento"
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
      Height          =   615
      Index           =   8
      Left            =   8235
      TabIndex        =   22
      Top             =   2790
      Width           =   1350
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Novo ToolTip:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   4590
      Width           =   990
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Novo Caption:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   4275
      Width           =   1035
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   570
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   3525
      Width           =   615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Objectos/Controles:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3960
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Salvar como:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   645
      Width           =   930
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   520
      X2              =   520
      Y1              =   16
      Y2              =   322
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   522
      X2              =   522
      Y1              =   18
      Y2              =   322
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "VB - Projeto:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   285
      Width           =   915
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gerador de Pacotes Multi-Linguagens"
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
      Height          =   1005
      Index           =   11
      Left            =   7950
      TabIndex        =   24
      Top             =   1350
      Width           =   1995
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004 Insígnia Desenvolvimento"
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
      Height          =   615
      Index           =   9
      Left            =   8235
      TabIndex        =   25
      Top             =   2805
      Width           =   1350
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
  frmAbout.Show 1
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

Private Sub cmdGenerate_Click()
  
  On Local Error Resume Next
  Dim i As Integer, j As Integer, ControlName As String
  
  ' Open the output file and generate the language pack
  Open sOutputFile For Output As #1
  
    ' Please, keep this comment lines here if you
    ' use this generator, thanks!
    Print #1, "; // Generated by Language Pack Generator Version " & App.Major & "." & App.Minor & App.Revision
    Print #1, "; // Developed by Frederico Machado (indiofu@bol.com.br)"
    Print #1, "; /////////////////////////////////////////////////////////"
    Print #1, ""
  
    For i = 1 To UBound(FormProp)
      Print #1, "[" & FormProp(i).Name & "]"
      Print #1, "Caption=" & Chr$(34) & IIf(Len(FormProp(i).NewCaption), FormProp(i).NewCaption, FormProp(i).Caption) & Chr$(34)
      For j = 1 To FormProp(i).ObjectNumber
        ControlName = FormProp(i).objProp(j).Name
        Print #1, FormProp(i).objProp(j).Name & ".Caption=" & Chr$(34) & IIf(Len(FormProp(i).objProp(j).NewCaption), FormProp(i).objProp(j).NewCaption, FormProp(i).objProp(j).Caption) & Chr$(34)
        If FormProp(i).objProp(j).NewToolTip <> "" Then
          Print #1, FormProp(i).objProp(j).Name & ".ToolTip=" & Chr$(34) & IIf(Len(FormProp(i).objProp(j).NewToolTip), FormProp(i).objProp(j).NewToolTip, FormProp(i).objProp(j).ToolTip) & Chr$(34)
        End If
      Next
      Print #1, "[" & FormProp(i).Name & ".End]"
      If i < UBound(FormProp) Then Print #1, ""
    Next
  
  Close #1
  
  If Err Then
    MsgBox "There was an error while creating the Language Pack.", vbCritical, "Erro"
    Kill sOutputFile
    Exit Sub
  End If
  
  MsgBox "Pacote criado com sucesso!", vbInformation
  
End Sub

Private Sub cmdHelp_Click()
  
  Dim sTmp As String
  
  sTmp = "To generate a Language Pack you need to:" & vbCrLf & vbCrLf
  sTmp = sTmp & "Select a VB Project, then click in the Load button. You can especify "
  sTmp = sTmp & "the Output File, but it can be especified before you click in the "
  sTmp = sTmp & "Generate Pack button. Select a form, select the controls and translate each "
  sTmp = sTmp & "one with your choosen language. Click in the Generate Pack button and "
  sTmp = sTmp & "your Language Pack is created and ready for use." & vbCrLf
  
  MsgBox sTmp, vbInformation, "Help"
  
End Sub

Private Sub cmdLoad_Click()
  
  Dim sLine As String, iCount As Integer
  
  If sVBProject = "" Then
    MsgBox "You need to select a VB Project to load forms and controls.", vbCritical
    txtProject.SetFocus
    Exit Sub
  End If
  
  Me.Caption = App.ProductName & " - Carregando ..."

  Me.MousePointer = 11
  tvForms.Nodes.Clear
  tvObjects.Nodes.Clear

  Dim sForms As String, sFArray() As String

  ' Scan the project for forms
  Open sVBProject For Input As #1
    Do
      Input #1, sLine
      If LCase(Left(sLine, 4)) = "name" Then
        tvForms.Nodes.Add , , "Project", Replace(Mid$(sLine, 6), Chr$(34), ""), 1
      ElseIf LCase(Left(sLine, 4)) = "form" Then
        sForms = sForms & Mid$(sLine, 6) & "|"
        iCount = iCount + 1
      End If
    Loop Until EOF(1)
  Close #1
  
  sFArray = Split(sForms, "|")
  tvForms.Nodes.Add "Project", tvwChild, "Forms", "Forms", 2
    
  ReDim FormProp(iCount)
  
  Dim i As Integer
  ' Get the properties and objects of each form
  For i = 1 To iCount
    tvForms.Nodes.Add "Forms", tvwChild, sFArray(i - 1), sFArray(i - 1), 4
    ReadFormFile sPrjFolder & sFArray(i - 1), i
  Next
  
  tvForms.Nodes.Item(1).Expanded = True
  tvForms.Nodes.Item(2).Expanded = True
  
  Me.MousePointer = 0
  
  Me.Caption = App.ProductName

End Sub

Sub CleanTexts()
  txtCaption = "": txtToolTip = "": txtNewCaption = "": txtNewToolTip = ""
End Sub

Private Sub Form_Load()
  Me.Caption = App.ProductName
  
  txtCaption.Enabled = False
  txtNewCaption.Enabled = False
  txtToolTip.Enabled = False
  txtNewToolTip.Enabled = False
  txtCaption.BackColor = &H8000000F
  txtNewCaption.BackColor = &H8000000F
  txtToolTip.BackColor = &H8000000F
  txtNewToolTip.BackColor = &H8000000F
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

Private Sub tvForms_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim ControlName As String
  
  If Node.Index < 3 Then Exit Sub
  tvObjects.Nodes.Clear
  CleanTexts
  
  iCurrForm = Node.Index - 2
  tvObjects.Nodes.Add , , "Form", FormProp(iCurrForm).Name
  
  ' List the objects and its properties in the list
  Dim i As Integer, bAdded(6) As Boolean
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
