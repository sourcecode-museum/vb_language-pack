VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmGrid 
   BackColor       =   &H00CC8A66&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grid View"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGrid 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBotoes 
      BackColor       =   &H00CDBBB6&
      Caption         =   "&Salvar"
      Height          =   765
      Index           =   0
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmGrid.frx":08E2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5625
      Width           =   1425
   End
   Begin VB.CommandButton cmdBotoes 
      BackColor       =   &H00CDBBB6&
      Caption         =   "&Cancelar"
      Height          =   765
      Index           =   1
      Left            =   9375
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmGrid.frx":0BEC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5625
      Width           =   1425
   End
   Begin MSFlexGridLib.MSFlexGrid flgGrid 
      Height          =   5520
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   9737
      _Version        =   393216
      RowHeightMin    =   300
      BackColorFixed  =   13482934
      BackColorBkg    =   15261661
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum eTipoGrid
    tgForms = 0
    tgStrings = 1
End Enum
Private pv_eTipoGrid As eTipoGrid
 
Private Const cHeadFormat = ""

Public Sub Carregar(ByVal pTipoGrid As eTipoGrid)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Select Case pTipoGrid
    Case Is = tgForms
        Me.Caption = "Grade Editor - " & FormProp(iCurrForm).Name
            
        flgGrid.Clear
        flgGrid.FormatString = "|Caption|Novo Caption|ToolTipText|Novo ToolipText"
        
        flgGrid.Rows = FormProp(iCurrForm).ObjectNumber + 1
        If flgGrid.Rows = 1 Then
            Unload Me
            Exit Sub
        Else
            With flgGrid
                .ColWidth(0) = 1500
                .ColWidth(1) = 2500
                .ColWidth(2) = 2500
                .ColWidth(3) = 2500
                .ColWidth(4) = 2500
            End With
        
            j = 1
            flgGrid.Row = 1
            For i = 1 To FormProp(iCurrForm).ObjectNumber
                If FormProp(iCurrForm).objProp(i).Type = "Menu" And FormProp(iCurrForm).objProp(i).Caption = "-" Then
                    k = k + 1
                Else
                    flgGrid.RowData(j) = i
                    flgGrid.Col = 0
                    flgGrid.Text = FormProp(iCurrForm).objProp(i).Name
                    flgGrid.Col = 1
                    flgGrid.Text = FormProp(iCurrForm).objProp(i).Caption
                    flgGrid.Col = 2
                    flgGrid.Text = FormProp(iCurrForm).objProp(i).NewCaption
                    flgGrid.Col = 3
                    flgGrid.Text = FormProp(iCurrForm).objProp(i).ToolTip
                    flgGrid.Col = 4
                    flgGrid.Text = FormProp(iCurrForm).objProp(i).NewToolTip
                    j = j + 1
                    If j < flgGrid.Rows Then flgGrid.Row = j
                End If
            Next
        
            flgGrid.Rows = flgGrid.Rows - k
        End If
        
        
    Case Is = tgStrings
        Caption = "Grade Editor - Strings"
        
        flgGrid.Clear
        flgGrid.FormatString = "|String"
        
        With flgGrid
            .ColWidth(0) = 1500
            .ColWidth(1) = 9000
        End With
    
        j = 1
        flgGrid.Rows = iStrings + 1
    
        flgGrid.Row = 1
        For i = 1 To iStrings
            flgGrid.RowData(j) = i
            flgGrid.Col = 0
            flgGrid.Text = LPGStrings(i).Name
            flgGrid.Col = 1
            flgGrid.Text = LPGStrings(i).String
            j = j + 1
            If j < flgGrid.Rows Then flgGrid.Row = j
        Next
        
    Case Else
        Unload Me
        Exit Sub
    End Select
    
    Me.Show vbModal
End Sub

Private Sub Recebe_Texto()
    With flgGrid
        txtGrid.Top = (.CellTop / 15) + .Top + 3
        txtGrid.Left = (.CellLeft / 15) + .Left + 3
        
        txtGrid.Width = .CellWidth / 15 - 3
        txtGrid.Height = 16
        txtGrid.Text = flgGrid.Text
        txtGrid.Visible = True
        txtGrid.SelStart = 0
        txtGrid.SelLength = Len(txtGrid.Text)

        txtGrid.SetFocus
    End With
End Sub

Private Sub cmdBotoes_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
    Case Is = 0 'Salvar
        If pv_eTipoGrid = tgStrings Then
            For i = 1 To flgGrid.Rows - 1
                flgGrid.Row = i
                If flgGrid.RowData(i) > 0 Then
                    flgGrid.Col = 1
                    LPGStrings(flgGrid.RowData(i)).String = flgGrid.Text
                End If
            Next
        
        Else
            For i = 1 To flgGrid.Rows - 1
                flgGrid.Row = i
                If flgGrid.RowData(i) > 0 Then
                    flgGrid.Col = 1
                    FormProp(iCurrForm).objProp(flgGrid.RowData(i)).Caption = flgGrid.Text
                    flgGrid.Col = 2
                    FormProp(iCurrForm).objProp(flgGrid.RowData(i)).NewCaption = flgGrid.Text
                    flgGrid.Col = 3
                    FormProp(iCurrForm).objProp(flgGrid.RowData(i)).ToolTip = flgGrid.Text
                    flgGrid.Col = 4
                    FormProp(iCurrForm).objProp(flgGrid.RowData(i)).NewToolTip = flgGrid.Text
                End If
            Next
        End If
        
    Case Is = 1 'Cancelar
        Unload Me
    End Select
End Sub

Private Sub flgGrid_Click()
    If flgGrid.Rows = 1 Then Exit Sub
    Call Recebe_Texto
End Sub

Private Sub flgGrid_KeyPress(KeyAscii As Integer)
    On Local Error GoTo ERRO
    With flgGrid
        Select Case KeyAscii
            Case vbKeyReturn
                If .Col = .Cols - 1 Then
                    .Row = .Row + 1
                    .Col = 1
                Else
                    .Col = .Col + 1
                End If
            Case vbKeyBack
                If Trim(.Text) <> "" Then
                    .Text = Mid(.Text, 1, Len(.Text) - 1)
                End If
            Case Is < 32
            Case Else
                If .Col = 0 Or .Row = 0 Then
                    Exit Sub
                Else
                    .Text = .Text & Chr(KeyAscii)
                End If
        End Select
    End With
ERRO:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGrid = Nothing
End Sub

Private Sub txtGrid_KeyPress(KeyAscii As Integer)
    On Local Error GoTo ERRO
    If KeyAscii = 13 Then
        flgGrid.Text = txtGrid.Text
        txtGrid.Text = ""
        txtGrid.Visible = False
        If flgGrid.Col = flgGrid.Cols - 1 Then
            flgGrid.Row = flgGrid.Row + 1
            flgGrid.Col = 0
        Else
            flgGrid.Col = flgGrid.Col + 1
        End If
    End If
ERRO:
    
End Sub

Private Sub txtGrid_LostFocus()
    txtGrid.Text = ""
    txtGrid.Visible = False
End Sub
