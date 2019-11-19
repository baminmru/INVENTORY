Attribute VB_Name = "modWizard"
Option Explicit



Function GetResString(nRes As Integer) As String
    Dim sTmp As String
    Dim sRetStr As String
  
    Do
        sTmp = LoadResString(nRes)
        If Right(sTmp, 1) = "_" Then
            sRetStr = sRetStr + VBA.Left(sTmp, Len(sTmp) - 1)
        Else
            sRetStr = sRetStr + sTmp
        End If
        nRes = nRes + 1
    Loop Until Right(sTmp, 1) <> "_"
    GetResString = sRetStr
  
End Function

Function GetField(sBuffer As String, sSep As String) As String
    Dim p As Integer
    
    p = InStr(sBuffer & sSep, sSep)
    GetField = VBA.Left(sBuffer, p - 1)
    sBuffer = Mid(sBuffer, p + Len(sSep))
  
End Function

Sub LoadResStrings(frm As Form)
    On Error Resume Next
    
    Dim ctl As Control
    Dim obj As Object
    
    'set the form's caption
    If IsNumeric(frm.Tag) Then
        frm.Caption = LoadResString(CInt(frm.Tag))
    End If
    
    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        If typename(ctl) = "Menu" Then
            If IsNumeric(ctl.Caption) Then
                If Err = 0 Then
                    ctl.Caption = LoadResString(CInt(ctl.Caption))
                Else
                    Err = 0
                End If
            End If
        ElseIf typename(ctl) = "TabStrip" Then
            For Each obj In ctl.Tabs
                If IsNumeric(obj.Tag) Then
                    obj.Caption = LoadResString(CInt(obj.Tag))
                End If
                'check for a tooltip
                If IsNumeric(obj.ToolTipText) Then
                    If Err = 0 Then
                        obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
                    Else
                        Err = 0
                    End If
                End If
            Next
        ElseIf typename(ctl) = "Toolbar" Then
            For Each obj In ctl.Buttons
                If IsNumeric(obj.Tag) Then
                    obj.ToolTipText = LoadResString(CInt(obj.Tag))
                End If
            Next
        ElseIf typename(ctl) = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                If IsNumeric(obj.Tag) Then
                    obj.Text = LoadResString(CInt(obj.Tag))
                End If
            Next
        Else
            If IsNumeric(ctl.Tag) Then
                If Err = 0 Then
                    ctl.Caption = GetResString(CInt(ctl.Tag))
                Else
                    Err = 0
                End If
            End If
            'check for a tooltip
            If IsNumeric(ctl.ToolTipText) Then
                If Err = 0 Then
                    ctl.ToolTipText = LoadResString(CInt(ctl.ToolTipText))
                Else
                    Err = 0
                End If
            End If
        End If
    Next

End Sub

'==================================================
'Purpose: Replace the <TOPIC_TEXT> string(s) in
'         res file string for correct placement
'         of localized tokens
'
'Inputs:  sString = String to search and replace in
'         sReplacement = String to replace token with
'         sReplacement2 = 2nd String to replace token with
'
'Outputs: New string with token replaced throughout
'==================================================
Function ReplaceTopicTokens(sString As String, _
                            sReplacement As String, _
                            sReplacement2 As String) As String
    On Error Resume Next
    
    Dim p As Integer
    Dim sTmp As String
    
    Const TOPIC_TEXT = "<TOPIC_TEXT>"
    Const TOPIC_TEXT2 = "<TOPIC_TEXT2>"
    
    sTmp = sString
    Do
        p = InStr(sTmp, TOPIC_TEXT)
        If p Then
            sTmp = VBA.Left(sTmp, p - 1) + sReplacement + Mid(sTmp, p + Len(TOPIC_TEXT))
        End If
    Loop While p
    
    If Len(sReplacement2) > 0 Then
        Do
            p = InStr(sTmp, TOPIC_TEXT2)
            If p Then
                sTmp = VBA.Left(sTmp, p - 1) + sReplacement2 + Mid(sTmp, p + Len(TOPIC_TEXT2))
            End If
        Loop While p
    End If
    
    ReplaceTopicTokens = sTmp
  
End Function

Public Function GetResData(sResName As String, sResType As String) As String
    Dim sTemp As String
    Dim p As Integer
  
    sTemp = StrConv(LoadResData(sResName, sResType), vbUnicode)
    p = InStr(sTemp, vbNullChar)
    If p Then sTemp = VBA.Left$(sTemp, p - 1)
    GetResData = sTemp
End Function





