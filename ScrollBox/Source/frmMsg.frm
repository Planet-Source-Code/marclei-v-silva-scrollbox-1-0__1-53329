VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4725
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1200
      Top             =   1800
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   570
      Top             =   1860
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   180
      Width           =   2025
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prompt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   150
      TabIndex        =   1
      Top             =   540
      Width           =   1965
      WordWrap        =   -1  'True
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   1350
      Width           =   105
   End
   Begin VB.Image imgMsg 
      Height          =   1800
      Left            =   0
      Picture         =   "frmMsg.frx":0442
      Top             =   0
      Width           =   2220
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frmMsg
'    Project    : CSScrollBox
'    Created By : Project Administrator
'    Description: Scroll box dialog
'
'    Modified   : 22/4/2004 18:47:57
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Private Const XGAP = 35

Private xTop As Long
Private xBottom As Long

Private m_ButtonIndex As Integer
Private m_TimeOut As Integer
Private m_bScrolling As Boolean
Private m_Left As Long
Private m_Top As Long
Private m_State As Integer
Private m_Action As Integer
Private m_Sound As ScrollBoxSound
Private m_Style As ScrollBoxStyle
Private m_Width As Integer
Private m_Buttons As Variant
Private m_Fnt As Variant
Private m_Title As String
Private m_ForeColor As Variant
Private m_Prompt As String
Private m_ButtonColor As Variant

Public Function Execute( _
    Prompt As Variant, _
    Optional nSecs As Variant = 5, _
    Optional Buttons As Variant, _
    Optional Title As Variant, _
    Optional Snd As ScrollBoxSound, _
    Optional Style As ScrollBoxStyle, _
    Optional Fnt As Variant, _
    Optional ForeColor As Variant, _
    Optional ButtonColor As Variant _
    ) As Long
    
    ' set form properties
    m_Sound = Snd
    m_Style = Style
    m_Buttons = Buttons
    m_Prompt = Prompt
    If IsNumeric(ForeColor) Then
        m_ForeColor = ForeColor
    End If
    If IsNumeric(ButtonColor) Then
        m_ButtonColor = ButtonColor
    End If
    If VarType(Title) = vbString Then
        m_Title = Title
    End If
    If IsObject(Fnt) Then
        Set m_Fnt = Fnt
    End If
    ' configure form timeout
    tmrTimeOut.Interval = nSecs * 1000
    ' enable initial scrolling
    tmrScroll.Enabled = True
    ' show the form
    Show vbModal
    ' returns user selection if buttons was passed
    Execute = m_Action
End Function

Private Sub Form_Load()
    ' prepare user interface
    PrepareUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project      :       CSScrollBox
' Procedure    :       PrepareUI
' Description  :       Prepare user interface
' Created by   :       Project Administrator
' Machine      :       ZEUS
' Date-Time    :       22/4/2004-18:47:39
'
' Parameters   :
' Return Values:
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub PrepareUI()
    Dim cur As Variant
    Dim TaskBarhWnd As Long
    Dim R As RECT
    Dim I As Integer
    Dim ab() As Byte
    Dim XOffset As Long
    Dim wid As Long
    
    m_Action = -1
    m_ButtonIndex = -1
    
    ' load cursor
    Set cur = LoadResPicture(101, vbResCursor)
    
    ' configure background image
    imgMsg.Left = 0
    imgMsg.Top = 0
    imgMsg.MousePointer = 99
    imgMsg.MouseIcon = cur
    
    ' load buttons if buttons was passed
    If IsArray(m_Buttons) Then
        For I = LBound(m_Buttons) To UBound(m_Buttons)
            ' load a button if not the main button
            If I <> 0 Then
                Load Button(I)
            End If
            Button(I).MousePointer = 99
            Button(I).MouseIcon = cur
            Button(I).Caption = " " & m_Buttons(I) & " "
            Button(I).Visible = True
            If IsEmpty(m_ButtonColor) = False And IsNumeric(m_ButtonColor) Then
                Button(I).BackColor = m_ButtonColor
            End If
            ' button label must be on top of the image control
            Button(I).ZOrder
            ' width to display the buttons
            If wid > 0 Then
                wid = wid + XGAP
            End If
            wid = wid + Button(I).Width
        Next
        ' configure buttons position
        XOffset = (imgMsg.Width - wid) / 2
        For I = Button.LBound To Button.UBound
            Button(I).Left = XOffset
            Button(I).Top = Button(0).Top
            XOffset = XOffset + Button(I).Width + XGAP
        Next
    Else
        Button(0).Visible = False
    End If
    
    ' set scrollbox title
    lblTitle.Caption = m_Title
    
    ' configure prompt
    lblPrompt.Caption = m_Prompt
    lblPrompt.MousePointer = 99
    lblPrompt.MouseIcon = cur
    If IsObject(m_Fnt) Then
        Set lblPrompt.Font = m_Fnt
    End If
    
    ' set text fore color
    If IsEmpty(m_ForeColor) = False And IsNumeric(m_ForeColor) Then
        lblPrompt.ForeColor = m_ForeColor
        lblTitle.ForeColor = m_ForeColor
    End If
    
    ' load style background
    ab = LoadResData(m_Style, "BACKGROUNDS")
    Set imgMsg.Picture = LoadPictureBytes(ab)
    
    ' set form boundaries...
    ' the form will fit the loaded image dimensions
    Me.Width = imgMsg.Width
    Me.Height = imgMsg.Height
    
    ' disable timeout
    tmrTimeOut.Enabled = False
    
    ' get windows taskbar boundaries
    TaskBarhWnd = FindWindow("Shell_TrayWnd", "")
    GetWindowRect TaskBarhWnd, R
    
    ' set up and top value
    ' in order to scroll the form
    xBottom = Screen.Height + Me.Height
    xTop = ScaleY(R.Top, vbPixels, vbTwips) - Me.Height
    m_Top = xBottom
    m_Left = Screen.Width - Me.Width

    ' set form initial position
    Me.Left = m_Left
    Me.Top = m_Top
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' rolagem para baixo
    ScrollDown
End Sub

Private Sub Button_MouseMove(Index As Integer, Btn As Integer, Shift As Integer, X As Single, Y As Single)
    If m_ButtonIndex <> Index Then
        If m_ButtonIndex <> -1 Then
            Button(m_ButtonIndex).BorderStyle = 0
        End If
        m_ButtonIndex = Index
        Button(Index).BorderStyle = 1
    End If
End Sub

Private Sub Button_Click(Index As Integer)
    m_Action = Index
    CloseForm
End Sub

Private Sub imgMsg_Click()
    m_Action = -1
    CloseForm
End Sub

Private Sub lblPrompt_Click()
    m_Action = -1
    CloseForm
End Sub

Private Sub CloseForm()
    ' disable timeout
    tmrTimeOut.Enabled = False
    ' unload window
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project      :       CSScrollBox
' Procedure    :       ScrollUp
' Description  :       Scroll the form up
' Created by   :       Project Administrator
' Machine      :       ZEUS
' Date-Time    :       22/4/2004-18:48:22
'
' Parameters   :
' Return Values:
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub ScrollUp()
    Dim lTop As Long
    Dim Snd() As Byte
    
    ' if still scrolling get out
    If m_bScrolling Then Exit Sub
    ' if a sound was specified, play it
    If m_Sound > 0 Then
        Snd = LoadResData(m_Sound, "WAV")
        PlayWaveBytes Snd, SND_ASYNC + SND_NOWAIT
    End If
    m_bScrolling = True
    StayOnTop hwnd, True
    DoEvents
    lTop = xBottom
    Do While lTop > xTop
        Wait 0.5            ' insert a delay here
        lTop = lTop - 30
        Me.Top = lTop
    Loop
    m_bScrolling = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project      :       CSScrollBox
' Procedure    :       ScrollDown
' Description  :       Scroll the form down
' Created by   :       Project Administrator
' Machine      :       ZEUS
' Date-Time    :       22/4/2004-18:48:39
'
' Parameters   :
' Return Values:
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub ScrollDown()
    Dim lTop As Long
    
    ' if it still scrolling get out
    If m_bScrolling = True Then Exit Sub
    m_bScrolling = True
    StayOnTop hwnd, True
    DoEvents
    lTop = xTop
    Do While lTop < xBottom
        Wait 0.5
        lTop = lTop + 30
        Me.Top = lTop
    Loop
    m_bScrolling = False
End Sub

Private Sub tmrScroll_Timer()
    ' disable initial scrolling
    tmrScroll.Enabled = False
    ' scroll window up
    ScrollUp
    ' check delay for scrolling down
    If tmrTimeOut.Interval <> 0 Then
        ' enable form timeout
        tmrTimeOut.Enabled = True
    End If
End Sub

Private Sub tmrTimeOut_Timer()
    ' disable time out
    tmrTimeOut.Enabled = False
    ' unloads the form
    Unload Me
End Sub
'-- end code
