VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TScrollBox"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : TScrollBox
'    Project    : CSScrollBox
'    Created By : Project Administrator
'    Description: Scroll box global class
'
'    Modified   : 22/4/2004 18:50:02
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Public Enum ScrollBoxSound
    sbNewMail = 101
    sbNewAlert = 102
    sbOnline = 103
    sbType = 104
End Enum

Public Enum ScrollBoxStyle
    sbDefault = 1205
    sbMessenger = 1206
End Enum

Public Function ScrollBox( _
    Prompt As Variant, _
    Optional Secs As Variant, _
    Optional Buttons As Variant, _
    Optional Title As Variant, _
    Optional Snd As ScrollBoxSound, _
    Optional Style As ScrollBoxStyle = sbDefault, _
    Optional Fnt As Variant, _
    Optional TextColor As Variant, _
    Optional ButtonColor As Variant _
    ) As Long
    
    Dim F As New frmMsg
    Dim Btns As Variant
    
    If IsArray(Buttons) Then
        Btns = Buttons
    Else
        If VarType(Buttons) = vbString Then
            If Len(Buttons) > 0 Then
                Btns = Split(Buttons, ";")
            End If
        End If
    End If
    ScrollBox = F.Execute(Prompt, Secs, Btns, Title, Snd, Style, Fnt, TextColor, ButtonColor)
    Unload F
    Set F = Nothing
End Function

