VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Texturized Text"
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   8775
      Begin VB.CommandButton Command4 
         Caption         =   "Texture"
         Height          =   495
         Left            =   7560
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Fade"
         Height          =   495
         Left            =   7560
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Preview"
         Height          =   375
         Left            =   7560
         TabIndex        =   18
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1320
         ScaleHeight     =   1155
         ScaleWidth      =   6015
         TabIndex        =   16
         Top             =   1920
         Width           =   6075
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1230
         Left            =   1320
         Picture         =   "Textout1.frx":0000
         ScaleHeight     =   1170
         ScaleWidth      =   6015
         TabIndex        =   14
         Top             =   360
         Width           =   6075
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1230
         Left            =   1320
         ScaleHeight     =   1170
         ScaleWidth      =   6015
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   6075
      End
      Begin VB.Label Label2 
         Caption         =   "Texturized Text"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Texture :"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Font"
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3735
      Begin VB.CheckBox Check3 
         Caption         =   "Anti-Aliasing"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Text            =   "0"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Specific Weight :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Italic"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Strikeout"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Underline"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Textout1.frx":21B6
         Left            =   2520
         List            =   "Textout1.frx":21EA
         TabIndex        =   6
         Text            =   "15"
         Top             =   360
         Width           =   975
      End
      Begin VB.PictureBox picSampleFont 
         AutoRedraw      =   -1  'True
         Height          =   1335
         Left            =   240
         ScaleHeight     =   1275
         ScaleWidth      =   3195
         TabIndex        =   5
         Top             =   2040
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rotated Text"
      Height          =   3495
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   100
         Left            =   240
         Max             =   900
         SmallChange     =   10
         TabIndex        =   12
         Top             =   3000
         Value           =   200
         Width           =   2655
      End
      Begin VB.TextBox txtAngle 
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Text            =   "20"
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   2340
         Left            =   240
         ScaleHeight     =   2280
         ScaleWidth      =   4245
         TabIndex        =   2
         Top             =   480
         Width           =   4305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Preview"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private hMasterFont As Long


Private Sub Check1_Click(Index As Integer)
    ShowSample
End Sub

Private Sub Check2_Click()
    Text1.Enabled = Check2.Value
    ShowSample
End Sub

Private Sub Check3_Click()
    ShowSample
End Sub

Private Sub Combo1_Click()
    ShowSample
End Sub

Private Sub Combo2_Change()
    If LenB(Combo2.Text) <> 0 Then ShowSample
End Sub

Private Sub Combo2_Click()
    ShowSample
End Sub
Private Function CreateFont_(Optional Rotation As Boolean = False) As Long
On Error Resume Next
    Dim plf As LOGFONT, I As Long
    Dim FontName As String
    FontName = Trim$(Combo1.List(Combo1.ListIndex))
    FontName = FontName + String(32 - Len(FontName), 0)
    For I = 1 To 32
        plf.lfFaceName(I) = Asc(Mid$(FontName, I, 1))
    Next
    'Height
    plf.lfHeight = CLng(Combo2.Text)
    'Width
    If Text1.Enabled Then
        plf.lfWidth = CLng(Text1.Text)
    Else
        plf.lfWidth = 0
    End If
    'Bold ,Underline ,Strikeout ,Italic
    If Check1(0).Value Then plf.lfWeight = 700
    If Check1(1).Value Then plf.lfUnderline = 1
    If Check1(2).Value Then plf.lfStrikeOut = 1
    If Check1(3).Value Then plf.lfItalic = 1
    'Anti Aliasing
    If Check3.Value Then
        plf.lfQuality = ANTIALIASED_QUALITY
    Else
        plf.lfQuality = NONANTIALIASED_QUALITY
    End If
    If Rotation Then plf.lfEscapement = CLng(txtAngle.Text) * 10
    CreateFont_ = CreateFontIndirect(plf)
End Function
Private Sub Comm()
Dim Rc As RECT
'hfnt = CreateFont(10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
'hprevfnt = CreateFont(10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

Dim lngTextX As Long
Dim plf As LOGFONT
Dim FontNames As String
    'lstrcpy plf.lfFaceName, "EBfar2"
    
    FontNames = "Comic Sans MS"
    FontNames = FontNames + String(32 - Len(FontNames), 0)
    For I = 1 To 32
        plf.lfFaceName(I) = Asc(Mid$(FontNames, I, 1))
    
    Next
    plf.lfHeight = 15 'FW_NORMAL
    GetClientRect hwnd, Rc
    SetBkMode hdc, TRANSPARENT
    hFnt = CreateFontIndirect(plf)
    hfntprev = SelectObject(hdc, hFnt)
    lngTextX = 10
    
    For I = 0 To 25
        
        Me.ForeColor = RGB(I * 10, (I * 5), 255 - (I * 10))
        TextOut hdc, lngTextX, 100, ChrW$(65 + I), 1
        lngTextX = lngTextX + (Me.TextWidth(ChrW(65 + I)) / Screen.TwipsPerPixelX) + 5

    Next
        
    SelectObject hdc, hfntprev
    DeleteObject hFnt

End Sub

Private Sub Command1_Click()
    Picture1.Cls
    Dim hfntprev As Long, I As Long, hFnt As Long
    hMasterFont = CreateFont_(True)
    SetBkMode Picture1.hdc, TRANSPARENT
    hfntprev = SelectObject(Picture1.hdc, hMasterFont)
    TextOut Picture1.hdc, 0, 120, "ABCabc Sample Text", 18
    SelectObject hdc, hfntprev
    DeleteObject hMasterFont
    Picture1.Refresh
End Sub

Private Sub Command2_Click()
Dim I As Long
    For I = 0 To Picture2.Width Step Screen.TwipsPerPixelX
    Picture2.ForeColor = RGB(I \ 30, 0, 255 - (I \ 30))
    Picture2.Line (I, 0)-(I, Picture2.Height)
    Next
    Command3_Click
End Sub

Private Sub Command3_Click()
    Picture3.Cls
    Dim hfntprev As Long, I As Long, hFnt As Long
    Dim pxWidth As Long, pxHeight As Long
    pxWidth = Picture2.Width \ Screen.TwipsPerPixelX
    pxHeight = Picture2.Height \ Screen.TwipsPerPixelY
    
    hMasterFont = CreateFont_
    SetBkMode Picture3.hdc, TRANSPARENT
    hfntprev = SelectObject(Picture3.hdc, hMasterFont)
    
    TextOut Picture3.hdc, 0, 0, "This is a Texturized and fade Text.", 35
    BitBlt Picture4.hdc, 0, 0, pxWidth, pxHeight, Picture3.hdc, 0, 0, NOTSRCCOPY
    BitBlt Picture4.hdc, 0, 0, pxWidth, pxHeight, Picture2.hdc, 0, 0, SRCAND
    BitBlt Picture3.hdc, 0, 0, pxWidth, pxHeight, Picture4.hdc, 0, 0, SRCPAINT
    Picture3.Refresh
    
    SelectObject hdc, hfntprev
    DeleteObject hMasterFont
    Picture3.Refresh
End Sub

Private Sub Command4_Click()
    Set Picture2.Picture = Picture2.Picture
    Command3_Click
End Sub

Private Sub Form_Load()
    Dim I As Long
    With Combo1
        For I = 0 To Screen.FontCount - 1
            .AddItem Screen.Fonts(I)
        Next I
    End With
    Combo1.Text = "Arial"
    Combo2.Text = 30
    Command1_Click
    Command3_Click
    End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If hMasterFont <> 0 Then DeleteObject hMasterFont
End Sub

Private Sub HScroll1_Change()
    txtAngle.Text = HScroll1.Value \ 10
End Sub

Private Sub HScroll1_Scroll()
    txtAngle.Text = HScroll1.Value \ 10
End Sub
Private Sub ShowSample()
    picSampleFont.Cls
    Dim hfntprev As Long, I As Long
    
    SetBkMode picSampleFont.hdc, TRANSPARENT
    hMasterFont = CreateFont_
    hfntprev = SelectObject(picSampleFont.hdc, hMasterFont)
    TextOut picSampleFont.hdc, 0, 0, "ABCabc Sample Text", 18
    SelectObject hdc, hfntprev
    picSampleFont.Refresh
    DeleteObject hMasterFont
    
    Command1_Click
    Command3_Click
End Sub

Private Sub Text1_Change()
    ShowSample
End Sub

Private Sub txtAngle_Change()
    Command1_Click
End Sub
