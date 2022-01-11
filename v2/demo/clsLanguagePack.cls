VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLanguagePack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // Language Pack Class Module Version 1.30
' // Developed by Frederico Machado (fredisoft@terra.com.br)
' // Vote for me if you like it please!
' /////////////////////////////////////////////////////////

Option Explicit

' Just to save the current pack loaded
Public sCurrentFile As String

' It saves the properties of objects
Private Type ObjectProperties
  Type       As String
  Name       As String
  Caption    As String
  ToolTip    As String
  NewCaption As String
  NewToolTip As String
  Index      As Long
End Type

' It saves the properties of forms
Private Type FormProperties
  FileName     As String
  Name         As String
  Caption      As String
  NewCaption   As String
  ObjectNumber As Integer
  objProp()    As ObjectProperties
End Type

' The variable that saves the properties and the variable that saves the number of forms
Private FormProp() As FormProperties
Private iFormCount As Integer
Private iStrings As Integer

' Strings
Private Type tStrings
  Name As String
  String As String
End Type
Private LPGStrings() As tStrings

Dim sVer As String * 4
Dim sProj As String * 32
Dim iForms As Integer

Function Decrypt(sText As String) As String
    Dim i As Integer, sT2 As String
    
    For i = 1 To Len(sText)
        sT2 = sT2 & Chr(Asc(Mid(sText, i, 1)) - 5)
    Next
    
    Decrypt = sT2
End Function

Sub LoadPack(sFile As String)
    Dim i As Integer, j As Integer
    
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
    
    sCurrentFile = sFile
End Sub

' As the name said, set the language pack in the form
Sub SetLanguageInForm(frmForm As Form)
  
  On Local Error Resume Next
  Dim i As Integer, j As Integer
  Dim iForm As Integer, ControlName As String, Index As Long
  
  ' It gets the index of the choosen form
  For i = 1 To iForms
    If FormProp(i).Name = frmForm.Name Then
      iForm = i
      Exit For
    End If
  Next i
  
  ' Set the caption of the form
  frmForm.Caption = IIf(Len(FormProp(iForm).NewCaption), FormProp(iForm).NewCaption, FormProp(iForm).Caption)
  ' Set the caption and tooltiptext of each control
  For j = 1 To FormProp(iForm).ObjectNumber
    If InStr(UCase$(FormProp(iForm).objProp(j).Name), ".TABCAPTION(") Then
        ControlName = Left$(FormProp(iForm).objProp(j).Name, InStr(FormProp(iForm).objProp(j).Name, ".") - 1)
        Index = Val(Mid$(FormProp(iForm).objProp(j).Name, InStrRev(FormProp(iForm).objProp(j).Name, "(") + 1))
        frmForm.Controls(ControlName).TabCaption(Index) = IIf(Len(FormProp(iForm).objProp(j).NewCaption), FormProp(iForm).objProp(j).NewCaption, FormProp(iForm).objProp(j).Caption)
    ElseIf InStr(FormProp(iForm).objProp(j).Name, ")") Then
        ControlName = Left$(FormProp(iForm).objProp(j).Name, InStr(FormProp(iForm).objProp(j).Name, "(") - 1)
        Index = Val(Mid$(FormProp(iForm).objProp(j).Name, InStrRev(FormProp(iForm).objProp(j).Name, "(") + 1))
        frmForm.Controls(ControlName)(Index).Caption = IIf(Len(FormProp(iForm).objProp(j).NewCaption), FormProp(iForm).objProp(j).NewCaption, FormProp(iForm).objProp(j).Caption)
    Else
        If Len(FormProp(iForm).objProp(j).Caption) Then frmForm.Controls(FormProp(iForm).objProp(j).Name).Caption = IIf(Len(FormProp(iForm).objProp(j).NewCaption), FormProp(iForm).objProp(j).NewCaption, FormProp(iForm).objProp(j).Caption)
        If Len(FormProp(iForm).objProp(j).ToolTip) Then frmForm.Controls(FormProp(iForm).objProp(j).Name).ToolTipText = IIf(Len(FormProp(iForm).objProp(j).NewToolTip), FormProp(iForm).objProp(j).NewToolTip, FormProp(iForm).objProp(j).ToolTip)
    End If
  Next j
  
End Sub

Function GetString(sName As String) As String
    Dim i As Integer
    For i = 1 To iStrings
        If LPGStrings(i).Name = sName Then
            GetString = LPGStrings(i).String
            Exit For
        End If
    Next
End Function

' Enumerate Language packs in the choosen folder
Function EnumLanguagePacks(sFolder As String, sExtension As String) As String

  ' Verify if the folder exists
  If Not DirExists(sFolder) Then
    MsgBox sFolder & " doesn't exist.", vbCritical
    Exit Function
  End If
  
  Dim sTmp As String
  
  ' Scan for language pack files
  If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
  sTmp = Dir$(sFolder & sExtension)
  If sTmp <> "" Then
    EnumLanguagePacks = sTmp
    sTmp = Dir$
    While Len(sTmp) > 0
      EnumLanguagePacks = EnumLanguagePacks & "|" & sTmp
      DoEvents
      sTmp = Dir$
    Wend
  End If

End Function

' This function verify if the choosen dir exists
' Returns True if the dir exists and False if it doesn't exist
Private Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    If Right$(strDirName, 1) <> "\" Then strDirName = strDirName & "\"
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)

    Err = 0
End Function

' ////////////////////////////////////////////////////////
' // Same subs and function, but in Portuguese (Brazil) //
' ////////////////////////////////////////////////////////

Sub CarregaPacotedeLinguagem(sArquivo As String)
  
  LoadPack sArquivo
  
End Sub

Sub SetaLinguagemnoForm(frmForm As Form)

  SetLanguageInForm frmForm

End Sub

Function EnumeraPacotesdeLinguagem(sDiretorio As String, sExtensao As String) As String
  
  EnumLanguagePacks sDiretorio, sExtensao
  
End Function
