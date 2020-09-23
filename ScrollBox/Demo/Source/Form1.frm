VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ScrollBox Demo"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Template 
      BackColor       =   &H80000018&
      Height          =   945
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4920
      Width           =   6795
   End
   Begin VB.ComboBox cboBtnColor 
      Height          =   315
      Left            =   3540
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2550
      Width           =   1065
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7110
      TabIndex        =   24
      Top             =   4890
      Width           =   1185
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   8460
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   8460
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":0000
         Height          =   555
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   8205
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ScrollBox Demo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   36
         Top             =   90
         Width           =   2445
      End
   End
   Begin VB.CommandButton cmdRevisions 
      Caption         =   "Revisions"
      Height          =   345
      Left            =   7110
      TabIndex        =   22
      Top             =   2040
      Width           =   1185
   End
   Begin VB.CommandButton cmdReadme 
      Caption         =   "Read me"
      Height          =   315
      Left            =   7110
      TabIndex        =   21
      Top             =   1650
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      Caption         =   "Font"
      Height          =   1545
      Left            =   90
      TabIndex        =   35
      Top             =   2970
      Width           =   3315
      Begin VB.ComboBox cboColor 
         Height          =   315
         Left            =   1260
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1020
         Width           =   1905
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Underline"
         Height          =   195
         Left            =   1770
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "Italic"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   1125
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "&Bold"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   825
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   2310
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2085
      End
      Begin VB.Label Label8 
         Caption         =   "Text color:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   1875
      End
   End
   Begin VB.TextBox Title 
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Text            =   "ScrollBox Demo"
      Top             =   1260
      Width           =   3315
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "<<"
      Height          =   315
      Index           =   1
      Left            =   4680
      TabIndex        =   11
      Top             =   2220
      Width           =   435
   End
   Begin VB.ListBox Buttons 
      Height          =   1035
      Left            =   5220
      TabIndex        =   12
      Top             =   1890
      Width           =   1635
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>"
      Height          =   315
      Index           =   0
      Left            =   4680
      TabIndex        =   10
      Top             =   1890
      Width           =   435
   End
   Begin VB.TextBox Button 
      Height          =   345
      Left            =   3540
      TabIndex        =   9
      Top             =   1890
      Width           =   1035
   End
   Begin VB.TextBox Result 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   6270
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4200
      Width           =   585
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7110
      TabIndex        =   23
      Top             =   2460
      Width           =   1185
   End
   Begin VB.TextBox Prompt 
      Height          =   975
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":008D
      Top             =   1890
      Width           =   3315
   End
   Begin VB.Frame Frame2 
      Caption         =   "Styles"
      Height          =   975
      Left            =   5280
      TabIndex        =   30
      Top             =   2970
      Width           =   1575
      Begin VB.OptionButton Style 
         Caption         =   "Messenger"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Tag             =   "1206"
         Top             =   570
         Width           =   1215
      End
      Begin VB.OptionButton Style 
         Caption         =   "Default"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Tag             =   "1205"
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox Interval 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5970
      TabIndex        =   8
      Text            =   "5"
      Top             =   1230
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sounds"
      Height          =   1575
      Left            =   3525
      TabIndex        =   28
      Top             =   2970
      Width           =   1635
      Begin VB.OptionButton snd 
         Caption         =   "New mail"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   330
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton snd 
         Caption         =   "New alert"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   615
         Width           =   1125
      End
      Begin VB.OptionButton snd 
         Caption         =   "On-line"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   885
         Width           =   1125
      End
      Begin VB.OptionButton snd 
         Caption         =   "Type"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Top             =   1170
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Height          =   375
      Left            =   7110
      TabIndex        =   20
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label Label10 
      Caption         =   "Button color:"
      Height          =   255
      Left            =   3540
      TabIndex        =   40
      Top             =   2280
      Width           =   945
   End
   Begin VB.Label Label9 
      Caption         =   "Scroll box template:"
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   4650
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Title:"
      Height          =   285
      Left            =   90
      TabIndex        =   34
      Top             =   1020
      Width           =   1125
   End
   Begin VB.Label Label4 
      Caption         =   "Butons:"
      Height          =   285
      Left            =   3540
      TabIndex        =   33
      Top             =   1650
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "Scroll box result:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5310
      TabIndex        =   32
      Top             =   4140
      Width           =   885
   End
   Begin VB.Label Label2 
      Caption         =   "Prompt:"
      Height          =   225
      Left            =   90
      TabIndex        =   31
      Top             =   1650
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Interval (secs). A zero value disable timeout:"
      Height          =   405
      Left            =   3570
      TabIndex        =   29
      Top             =   1200
      Width           =   2385
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : Form1
'    Project    : Project1
'    Created By : Project Administrator
'    Description: ScrollBox demo project form
'
'    Modified   : 22/4/2004 19:02:31
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Private Sub Button_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdButton_Click 0
        KeyCode = 0
    End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
    If Index = 0 Then
        If Len(Button.Text) > 0 Then
            If Buttons.ListCount > 4 Then
                If MsgBox("There are too many buttons. Some buttons may not be visible. Continue?", vbYesNo + vbExclamation) = vbNo Then
                    Exit Sub
                End If
            End If
            Buttons.AddItem Button.Text
            Buttons.ListIndex = Buttons.NewIndex
            Button.Text = ""
            Button.SetFocus
        End If
    Else
        If Buttons.ListCount > 0 Then
            If Buttons.SelCount > 0 Then
                Button.Text = Buttons.List(Buttons.ListIndex)
                Buttons.RemoveItem Buttons.ListIndex
                Buttons.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Template.Text
End Sub

Private Sub cmdRevisions_Click()
    OpenFile App.Path & "\revisions.rtf", "Revisions"
End Sub

Private Sub cmdShow_Click()
    Dim sn As ScrollBoxSound
    Dim st As ScrollBoxStyle
    Dim nRet As Long
    Dim Btns As Variant
    Dim Fnt As StdFont
    Dim Color As Long
    Dim sTemp As String
    Dim BtnCOlor As Long
    
    ' get user settings
    sn = GetSound           ' get box sound
    st = GetStyle           ' get box style
    Btns = GetButtons       ' get user defined buttons
    Set Fnt = GetFont       ' get dialog font
    Color = GetColor        ' get text color
    BtnCOlor = GetBtnColor  ' get button color
    
    ' show template
    sTemp = "Call ScrollBox(" & GetText(Prompt.Text) & ", " & Val(Interval.Text) & ", """ & Btns & """, """ & Title.Text & """, " & sn & ", " & st & ",, " & Color & ", " & BtnCOlor & ")"
    Template.Text = sTemp
    
    ' show scrollbox
    nRet = ScrollBox(Prompt.Text, Val(Interval), Btns, Title.Text, sn, st, Fnt, Color, BtnCOlor)
    
    ' register user return action
    Result.Text = nRet
End Sub

Private Function GetSound() As Integer
    Dim I As Integer
    
    For I = snd.LBound To snd.UBound
        If snd(I).Value = True Then
            GetSound = 101 + I
            Exit Function
        End If
    Next
End Function

Private Function GetStyle() As Integer
    Dim I As Integer
    
    For I = Style.LBound To Style.UBound
        If Style(I).Value = True Then
            GetStyle = Val(Style(I).Tag)
            Exit Function
        End If
    Next
End Function

Private Function GetButtons() As String
    Dim I As Integer
    Dim SrcStr As String
    
    For I = 0 To Buttons.ListCount - 1
        If Len(SrcStr) > 0 Then
            SrcStr = SrcStr & ";"
        End If
        SrcStr = SrcStr & Buttons.List(I)
    Next
    GetButtons = SrcStr
End Function

Private Function GetFont() As StdFont
    Dim Fnt As New StdFont
    
    With Fnt
        .Name = cboFont.Text
        .Size = Val(cboSize.Text)
        .Bold = (chkBold.Value = vbChecked)
        .Italic = (chkItalic.Value = vbChecked)
        .Underline = (chkUnderline.Value = vbChecked)
    End With
    Set GetFont = Fnt
End Function

Private Function GetColor() As Long
    If cboColor.ListIndex = -1 Then
        GetColor = vbBlack
    Else
        GetColor = cboColor.ItemData(cboColor.ListIndex)
    End If
End Function

Private Function GetBtnColor() As Long
    If cboBtnColor.ListIndex = -1 Then
        GetBtnColor = vbInfoBackground
    Else
        GetBtnColor = cboBtnColor.ItemData(cboBtnColor.ListIndex)
    End If
End Function

Private Sub cmdReadme_Click()
    OpenFile App.Path & "\readme.rtf", "Read me"
End Sub

Private Sub OpenFile(FileName As String, FileTitle As String)
    Dim f As New Form2
    
    f.Execute FileName, FileTitle
End Sub

Private Sub Form_Load()
    AddItem cboFont, "Tahoma"
    AddItem cboFont, "Arial"
    AddItem cboFont, "Courier New"
    AddItem cboFont, "Verdana"
    AddItem cboFont, "Times New Roman"
    cboFont.ListIndex = 2
    
    AddItem cboSize, "7"
    AddItem cboSize, "8"
    AddItem cboSize, "9"
    AddItem cboSize, "10"
    cboSize.ListIndex = 1
    
    AddItem cboColor, "Red", vbRed
    AddItem cboColor, "Blue", vbBlue
    AddItem cboColor, "Magenta", vbMagenta
    AddItem cboColor, "Black", vbBlack

    AddItem cboBtnColor, "ButtonFace", vbButtonFace
    AddItem cboBtnColor, "ToolTip", vbInfoBackground
    AddItem cboBtnColor, "Desktop", vbDesktop
    AddItem cboBtnColor, "White", vbWhite
End Sub

Private Sub AddItem(rList As ComboBox, Text As String, Optional Value As Long)
    rList.AddItem Text
    rList.ItemData(rList.NewIndex) = Value
End Sub

Private Function GetText(s As String) As String
    Dim arr As Variant
    Dim elem As Variant
    Dim sTemp As String
    
    arr = Split(s, vbCrLf)
    For Each elem In arr
        If Len(sTemp) > 0 Then
            sTemp = sTemp & " & vbCrLf & "
        End If
        sTemp = sTemp & Chr(34) & elem & Chr(34)
    Next
    GetText = sTemp
End Function
