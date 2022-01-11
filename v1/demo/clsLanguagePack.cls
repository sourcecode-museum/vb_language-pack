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
' // Language Pack Class Module Version 1.20
' // Developed by Frederico Machado (indiofu@bol.com.br)
' // Vote for me if you like it please!
' /////////////////////////////////////////////////////////

Option Explicit

' Just to save the current pack loaded
Public sCurrentFile As String

' It saves the properties of objects
Private Type ObjectProperties
  Name As String
  Caption As String
  ToolTip As String
End Type

' It saves the properties of forms
Private Type FormProperties
  Name As String
  Caption As String
  ObjectCount As Integer
  ObjProp() As ObjectProperties
End Type

' The variable that saves the properties and the variable that saves the number of forms
Private FormProp() As FormProperties
Private iFormCount As Integer

' It loads the entire language pack
Sub LoadLanguagePack(sFile As String)

  ' Just some variables
  Dim sLine As String, iPos As Integer, sTmp As String
  Dim sFormName As String, sTmp2 As String
  Dim bFormFound As Boolean

  ' Set the current pack used and set the nuber of forms to 0
  sCurrentFile = sFile
  iFormCount = 0

  ' Open the language pack file
  Open sFile For Input As #1
    Do
      ' Get a line
      Input #1, sLine
      ' If the line starts with ; it is a comment line
      ' If the line is a blank line then go to next line
      If Left$(sLine, 1) = ";" Or sLine = "" Then GoTo Jump
      ' End of form objects and properties
      If Left$(sLine, 1) = "[" And Right$(sLine, 5) = ".End]" Then
        bFormFound = False: GoTo Jump
      End If
      ' Begin of form objects and properties
      If Left$(sLine, 1) = "[" And Right$(sLine, 1) = "]" Then
        bFormFound = True
        sFormName = Mid$(sLine, 2, Len(sLine) - 2)
        iFormCount = iFormCount + 1
        ReDim Preserve FormProp(iFormCount)
        FormProp(iFormCount).Name = sFormName: GoTo Jump
      End If
      ' Form Caption found
      If Left$(sLine, 7) = "Caption" Then
        sTmp = Mid$(sLine, InStr(sLine, "=") + 2)
        FormProp(iFormCount).Caption = Left$(sTmp, Len(sTmp) - 1)
        GoTo Jump
      End If
      ' Verify if it's the caption properties of the object
      iPos = InStr(sLine, ".Caption")
      ' Caption was found
      If iPos > 0 And bFormFound Then
        FormProp(iFormCount).ObjectCount = FormProp(iFormCount).ObjectCount + 1
        ReDim Preserve FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount)
        sTmp = Left$(sLine, iPos - 1)
        FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount).Name = sTmp
        sTmp = Mid$(sLine, InStr(sLine, "=") + 2)
        ' It verifys if VB got the entire line
        ' The command Input #1, sLine gets a line
        ' but if it has a ',' then VB thinks that
        ' it is another line. Strange.
        ' (Chr$(34) = '"' (comma I think))
        If Right$(sTmp, 1) <> Chr$(34) Then
          Do While Right$(sTmp, 1) <> Chr$(34)
            Input #1, sTmp2
            sTmp = sTmp & ", " & sTmp2
          Loop
        End If
        If Right$(sTmp, 1) = Chr$(34) Then sTmp = Left$(sTmp, Len(sTmp) - 1)
        ' Set the propertie
        FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount).Caption = sTmp
        GoTo Jump
      End If
      ' Verify if it is a ToolTipText
      iPos = InStr(sLine, ".ToolTip")
      If iPos > 0 And bFormFound Then
        sTmp = Left$(sLine, iPos - 1)
        FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount).Name = sTmp
        sTmp = Mid$(sLine, InStr(sLine, "=") + 2)
        ' This is the same thing in the Caption
        ' propertie above.
        If Right$(sTmp, 1) <> Chr$(34) Then
          Do While Right$(sTmp, 1) <> Chr$(34)
            Input #1, sTmp2
            sTmp = sTmp & ", " & sTmp2
          Loop
        End If
        If Right$(sTmp, 1) = Chr$(34) Then sTmp = Left$(sTmp, Len(sTmp) - 1)
        ' Set the propertie
        FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount).ToolTip = sTmp
        GoTo Jump
      End If
      
Jump:

    ' Loop until End Of File
    Loop Until EOF(1)
  ' Close the pack
  Close #1

End Sub

' As the name said, set the language pack in the form
Sub SetLanguageInForm(frmForm As Form)
  
  On Local Error Resume Next
  Dim i As Integer, j As Integer
  Dim iForm As Integer, ControlName As String, Index As Long
  
  ' It gets the index of the choosen form
  For i = 1 To iFormCount
    If FormProp(i).Name = frmForm.Name Then
      iForm = i
      Exit For
    End If
  Next i
  
  ' Set the caption of the form
  frmForm.Caption = FormProp(iForm).Caption
  ' Set the caption and tooltiptext of each control
  For j = 1 To FormProp(iForm).ObjectCount
    If InStr(UCase$(FormProp(iForm).ObjProp(j).Name), ".TABCAPTION(") Then
        ControlName = Left$(FormProp(iForm).ObjProp(j).Name, InStr(FormProp(iForm).ObjProp(j).Name, ".") - 1)
        Index = Val(Mid$(FormProp(iForm).ObjProp(j).Name, InStrRev(FormProp(iForm).ObjProp(j).Name, "(") + 1))
        frmForm.Controls(ControlName).TabCaption(Index) = FormProp(iForm).ObjProp(j).Caption
    ElseIf InStr(FormProp(iForm).ObjProp(j).Name, ")") Then
        ControlName = Left$(FormProp(iForm).ObjProp(j).Name, InStr(FormProp(iForm).ObjProp(j).Name, "(") - 1)
        Index = Val(Mid$(FormProp(iForm).ObjProp(j).Name, InStrRev(FormProp(iForm).ObjProp(j).Name, "(") + 1))
        frmForm.Controls(ControlName)(Index).Caption = FormProp(iForm).ObjProp(j).Caption
    Else
        If Len(FormProp(iForm).ObjProp(j).Caption) Then frmForm.Controls(FormProp(iForm).ObjProp(j).Name).Caption = FormProp(iForm).ObjProp(j).Caption
        If Len(FormProp(iForm).ObjProp(j).ToolTip) Then frmForm.Controls(FormProp(iForm).ObjProp(j).Name).ToolTipText = FormProp(iForm).ObjProp(j).ToolTip
    End If
  Next j
  
End Sub

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
  
  LoadLanguagePack sArquivo
  
End Sub

Sub SetaLinguagemnoForm(frmForm As Form)

  SetLanguageInForm frmForm

End Sub

Function EnumeraPacotesdeLinguagem(sDiretorio As String, sExtensao As String) As String
  
  EnumLanguagePacks sDiretorio, sExtensao
  
End Function
