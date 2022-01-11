Attribute VB_Name = "modMain"
' //
' // VB Language Pack Generator 1.20
' // Developed by Frederico Machado (indiofu@bol.com.br)
' // Vote for me if you like it please!
' ////////////////////////////////////////////////////////

Option Explicit

Global sVBProject  As String
Global sPrjFolder  As String
Global sOutputFile As String

Global iCurrForm As Integer
Global iCurrObj  As Integer

Public Type ObjectProperties
  Type       As String
  Name       As String
  Caption    As String
  ToolTip    As String
  NewCaption As String
  NewToolTip As String
  Index      As Long
End Type

Public Type FormProperties
  Name         As String
  Caption      As String
  NewCaption   As String
  ObjectNumber As Integer
  objProp()    As ObjectProperties
End Type

Global FormProp() As FormProperties

Public Const sTObjects As String = "CommandButton Frame CheckBox Label OptionButton Menu SSTab"

' It reads the form file looking for objects and properties
Public Sub ReadFormFile(sFile As String, Index As Integer)

  Dim sLine As String, bForm As Boolean, iObject As Integer
  Dim iPos As Integer, sTmp As String, sTmp2 As String, SSTabObject As String
  Dim sFrxFile As String, lFrxPos As Long
  
  On Local Error Resume Next
  
  Open sFile For Input As #1
    Do
      Input #1, sLine
      If InStr(sLine, "Begin VB.Form") > 0 Or InStr(sLine, "Begin VB.MDIForm") > 0 Then
        FormProp(Index).Name = Replace(Mid$(sLine, 15), " ", "")
        bForm = True: GoTo Jump
      End If
      If sLine = "End" And bForm And iObject > 0 Then iObject = iObject - 1
      If InStr(sLine, "Caption") > 0 And iObject = 0 Then
        iPos = InStr(sLine, Chr$(34)) + 1
        FormProp(Index).Caption = Mid$(sLine, iPos, Len(sLine) - iPos)
        GoTo Jump
      End If
      iPos = InStr(sLine, "Begin VB.")
      If iPos > 0 Then
        iObject = iObject + 1
        iPos = iPos + 9
        sTmp = Mid$(sLine, iPos, InStr(iPos, sLine, " ") - iPos)
        If InStr(sTObjects, sTmp) > 0 Then
          FormProp(Index).ObjectNumber = FormProp(Index).ObjectNumber + 1
          ReDim Preserve FormProp(Index).objProp(FormProp(Index).ObjectNumber)
          FormProp(Index).objProp(FormProp(Index).ObjectNumber).Type = sTmp
          sTmp = Replace(Mid$(sLine, iPos + Len(sTmp) + 1), " ", "")
          FormProp(Index).objProp(FormProp(Index).ObjectNumber).Name = sTmp
        End If
        GoTo Jump
      End If
      
      iPos = InStr(sLine, "Begin TabDlg")
      If iPos > 0 Then
        iObject = iObject + 1
        iPos = iPos + 9
        sTmp = Mid$(sLine, iPos + 4, InStr(iPos, sLine, " ") - iPos - 4)
        If InStr(sTObjects, sTmp) > 0 Then
          'FormProp(Index).ObjectNumber = FormProp(Index).ObjectNumber + 1
          'ReDim Preserve FormProp(Index).objProp(FormProp(Index).ObjectNumber)
          'FormProp(Index).objProp(FormProp(Index).ObjectNumber).Type = sTmp
          sTmp = Replace(Mid$(sLine, iPos + Len(sTmp) + 5), " ", "")
          'FormProp(Index).objProp(FormProp(Index).ObjectNumber).Name = sTmp
          SSTabObject = sTmp
        End If
        GoTo Jump
      End If
      
      iPos = InStr(sLine, "Caption")
      If iPos > 0 Then
        If InStr(sLine, "$" & Chr$(34)) = 0 Then
          iPos = InStr(sLine, Chr$(34)) + 1
          sTmp = Mid$(sLine, iPos)
          If Right$(sTmp, 1) <> Chr$(34) Then
            Do While Right$(sTmp, 1) <> Chr$(34)
              Input #1, sTmp2
              sTmp = sTmp & ", " & sTmp2
            Loop
          End If
          If Right$(sTmp, 1) = Chr$(34) Then sTmp = Left$(sTmp, Len(sTmp) - 1)
          If InStr(sLine, "TabCaption") Then
            FormProp(Index).ObjectNumber = FormProp(Index).ObjectNumber + 1
            ReDim Preserve FormProp(Index).objProp(FormProp(Index).ObjectNumber)
            FormProp(Index).objProp(FormProp(Index).ObjectNumber).Name = SSTabObject & "." & Left$(sLine, InStr(sLine, " ") - 1)
            FormProp(Index).objProp(FormProp(Index).ObjectNumber).Caption = sTmp
            FormProp(Index).objProp(FormProp(Index).ObjectNumber).Type = "SSTab"
          Else
            FormProp(Index).objProp(FormProp(Index).ObjectNumber).Caption = sTmp
            If FormProp(Index).objProp(FormProp(Index).ObjectNumber).Type = "Menu" And FormProp(Index).objProp(FormProp(Index).ObjectNumber).Caption = "-" Then
              FormProp(Index).objProp(FormProp(Index).ObjectNumber).NewCaption = "-"
            End If
          End If
        Else
          iPos = InStr(sLine, Chr$(34)) + 1
          sFrxFile = Mid$(sLine, iPos, InStr(iPos, sLine, Chr$(34)) - iPos)
          iPos = InStrRev(sLine, ":")
          lFrxPos = "&H" & Right(sLine, Len(sLine) - iPos)
          FormProp(Index).objProp(FormProp(Index).ObjectNumber).Caption = GetPropertie(sFrxFile, lFrxPos)
        End If
        GoTo Jump
      End If
      iPos = InStr(sLine, "ToolTipText")
      If iPos > 0 Then
        If InStr(sLine, "$" & Chr$(34)) = 0 Then
          iPos = InStr(sLine, Chr$(34)) + 1
          sTmp = Mid$(sLine, iPos)
          If Right$(sTmp, 1) <> Chr$(34) Then
            MsgBox sTmp
            Do While Right$(sTmp, 1) <> Chr$(34)
              Input #1, sTmp2
              sTmp = sTmp & ", " & sTmp2
            Loop
          End If
          If Right$(sTmp, 1) = Chr$(34) Then sTmp = Left$(sTmp, Len(sTmp) - 1)
          FormProp(Index).objProp(FormProp(Index).ObjectNumber).ToolTip = sTmp
        Else
          iPos = InStr(sLine, Chr$(34)) + 1
          sFrxFile = Mid$(sLine, iPos, InStr(iPos, sLine, Chr$(34)) - iPos)
          iPos = InStrRev(sLine, ":")
          lFrxPos = "&H" & Right(sLine, Len(sLine) - iPos)
          FormProp(Index).objProp(FormProp(Index).ObjectNumber).ToolTip = GetPropertie(sFrxFile, lFrxPos)
        End If
        GoTo Jump
      End If
      iPos = InStr(sLine, "Index")
      If iPos = 1 Then
        iPos = InStr(sLine, "=") + 1
        sTmp = Trim$(Mid$(sLine, iPos))
        FormProp(Index).objProp(FormProp(Index).ObjectNumber).Index = sTmp
        FormProp(Index).objProp(FormProp(Index).ObjectNumber).Name = FormProp(Index).objProp(FormProp(Index).ObjectNumber).Name & "(" & sTmp & ")"
        GoTo Jump
      End If
      If Left$(sLine, 9) = "Attribute" And bForm And iObject = 0 Then Exit Do
      
Jump:
      
    Loop Until EOF(1)
  Close #1

End Sub

' If the caption or tooltip are too large, then
' VB puts it in the .FRX file (binary)
' This function reads the FRX file and returns the
' propertie.
Public Function GetPropertie(sFile As String, lPos As Long) As String

  Dim sReturn As String
  
  Open sFile For Binary As #2
    Get #2, lPos + 5, sReturn
    sReturn = Input(500, 2)
  Close #2
  
  lPos = InStr(sReturn, Chr$(0))
  If lPos > 0 Then
    GetPropertie = Left$(sReturn, lPos - 2)
  Else
    GetPropertie = sReturn
  End If

End Function
