VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9510
   LinkTopic       =   "Form2"
   ScaleHeight     =   7560
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   7035
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   12409
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form2.frx":0000
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "File"
      Begin VB.Menu mnuFile 
         Caption         =   "Save"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_FileName As String
Private m_Modified As Boolean

Public Sub Execute(FileName As String, Title As String)
    m_FileName = FileName
    Caption = Title
    rtfText.LoadFile FileName
    m_Modified = False
    Show vbModal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_Modified Then
        Dim nRet As VbMsgBoxResult
        
        nRet = MsgBox("File was modified. Do you want to save changes?", vbQuestion + vbYesNoCancel)
        Select Case nRet
            Case vbYes: Save
            Case vbCancel: Cancel = True
            Case vbNo: ' nothing
        End Select
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 50, 50, ScaleWidth - 100, ScaleHeight - 100
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
        Case 0: Save
        Case 2: Unload Me
    End Select
End Sub

Private Sub rtfText_Change()
    m_Modified = True
End Sub

Private Sub Save()
    rtfText.SaveFile m_FileName
    m_Modified = False
End Sub
