VERSION 5.00
Object = "{A9E16F1F-AC4C-11D4-BE93-9EDDB85F6233}#2.0#0"; "MessageBox.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Box Testing Console - Customize your own message box easily!"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "frmMsgBoxTest"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad2 
      Caption         =   "..."
      Height          =   315
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdLoad1 
      Caption         =   "..."
      Height          =   315
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   120
      Width           =   375
   End
   Begin MessageBox.Message MessageBox 
      Left            =   765
      Top             =   2460
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.TextBox txtSmallPic 
      Height          =   315
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "(None)"
      Top             =   480
      Width           =   2970
   End
   Begin VB.TextBox txtBox 
      Height          =   315
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "(None)"
      Top             =   120
      Width           =   2970
   End
   Begin VB.CommandButton Button1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Button 1"
      Default         =   -1  'True
      Height          =   375
      Left            =   7500
      MouseIcon       =   "frmTest.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Button2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Button 2"
      Height          =   375
      Left            =   8480
      MouseIcon       =   "frmTest.frx":0594
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Button3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Button 3"
      Height          =   375
      Left            =   9570
      MouseIcon       =   "frmTest.frx":06E6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Button4 
      BackColor       =   &H00FFC0FF&
      Cancel          =   -1  'True
      Caption         =   "Button 4"
      Height          =   375
      Left            =   10530
      MouseIcon       =   "frmTest.frx":0838
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox cboBStyle 
      Height          =   330
      ItemData        =   "frmTest.frx":098A
      Left            =   1080
      List            =   "frmTest.frx":0994
      TabIndex        =   1
      Text            =   "cboBStyle"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton minLowLight 
      Caption         =   "Non-Highlighted"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton minHighLight 
      Caption         =   "Highlighted button"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton minText 
      Caption         =   $"frmTest.frx":09B3
      Height          =   1095
      Left            =   3840
      TabIndex        =   15
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton minIcon 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton minTitle 
      Caption         =   "Click - Color, Right Click - Font"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   1800
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1440
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".bmp"
      DialogTitle     =   "Browse for Picture"
      Filter          =   $"frmTest.frx":0A8A
   End
   Begin VB.Timer timUpdate 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   2400
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C000&
      Caption         =   "E&xit Test Program"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4080
      Width           =   3375
   End
   Begin VB.ComboBox cboStyle 
      Height          =   330
      ItemData        =   "frmTest.frx":0B20
      Left            =   120
      List            =   "frmTest.frx":0B30
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtButton1 
      Height          =   315
      Left            =   3720
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtButton2 
      Height          =   315
      Left            =   5400
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtButton3 
      Height          =   315
      Left            =   3720
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtButton4 
      Height          =   315
      Left            =   5400
      TabIndex        =   10
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdHide 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Hide it!"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Show it!"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtMessage 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtTop 
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtLeft 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton minBox 
      Height          =   1695
      Left            =   3720
      TabIndex        =   14
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFC0&
      X1              =   5880
      X2              =   10680
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Shape shpCircle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00004000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   11160
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   255
   End
   Begin VB.Shape shpCover2 
      BorderColor     =   &H00004000&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   10800
      Top             =   3960
      Width           =   855
   End
   Begin VB.Shape shpCover 
      BorderColor     =   &H00000040&
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6960
      Top             =   3960
      Width           =   3855
   End
   Begin VB.Label Status 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*Note: Double-Click at Textbox to clear each picture."
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   7320
      TabIndex        =   42
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label lblVer 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Version 3.0 (C) Copyright 2000 All Right Reserved"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6960
      TabIndex        =   41
      Top             =   4200
      Width           =   4695
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   7320
      X2              =   8280
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   7800
      X2              =   8760
      Y1              =   720
      Y2              =   1560
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   7320
      X2              =   8280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblTextPic 
      BackColor       =   &H0000FF00&
      Caption         =   "Text Picture:"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   7320
      TabIndex        =   36
      Top             =   515
      Width           =   975
   End
   Begin VB.Label lblMainPic 
      BackColor       =   &H00FF0000&
      Caption         =   "*Box Picture:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   7320
      TabIndex        =   35
      Top             =   170
      Width           =   975
   End
   Begin VB.Image imgTextPic 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   7500
      Picture         =   "frmTest.frx":0B72
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Line Line8 
      X1              =   7200
      X2              =   7200
      Y1              =   0
      Y2              =   3840
   End
   Begin VB.Label lblCredits 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Submitted on PSC by Alpha (hyperactive0000@hotmail.com)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   20
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label lblMini 
      BackColor       =   &H0080C0FF&
      Caption         =   "Below is a Minimap of the Message Box:"
      Height          =   255
      Left            =   3720
      TabIndex        =   30
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label minClose 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderStyle     =   2  'Dash
      X1              =   3600
      X2              =   7200
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line6 
      X1              =   3600
      X2              =   3600
      Y1              =   0
      Y2              =   4560
   End
   Begin VB.Line Line5 
      BorderStyle     =   2  'Dash
      X1              =   3600
      X2              =   11760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   3600
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H0080C0FF&
      Caption         =   "Message Box Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button 1"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   25
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button 2"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   26
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button 3"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   27
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button 4"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   28
      Top             =   720
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   3600
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lblText 
      BackColor       =   &H0080C0FF&
      Caption         =   "Message Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   3600
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblTop 
      BackColor       =   &H0080C0FF&
      Caption         =   "Top:"
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblLeft 
      BackColor       =   &H0080C0FF&
      Caption         =   "Left:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   3600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblBorderStyle 
      BackColor       =   &H0080C0FF&
      Caption         =   "Border Style:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   145
      Width           =   975
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   7440
      X2              =   7800
      Y1              =   1200
      Y2              =   360
   End
   Begin VB.Image imgBox 
      BorderStyle     =   1  'Fixed Single
      Height          =   2475
      Left            =   7320
      Picture         =   "frmTest.frx":0BF8
      Stretch         =   -1  'True
      Top             =   960
      Width           =   4335
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button1_Click()
    MessageBox.SetButtonFocus 1
    MessageBox_Btn1Focus
    Status = "Status: Button 1 Focus Set"
End Sub

Private Sub Button2_Click()
    MessageBox.SetButtonFocus 2
    MessageBox_Btn2Focus
    Status = "Status: Button 2 Focus Set"
End Sub

Private Sub Button3_Click()
    MessageBox.SetButtonFocus 3
    MessageBox_Btn3Focus
    Status = "Status: Button 3 Focus Set"
End Sub

Private Sub Button4_Click()
    MessageBox.SetButtonFocus 4
    MessageBox_Btn4Focus
    Status = "Status: Button 4 Focus Set"
End Sub

Private Sub cboBStyle_Change()
    MessageBox.BorderStyle = cboBStyle.ListIndex
End Sub

Private Sub cboStyle_Click()
    MessageBox.MsgBoxStyle = cboStyle.ListIndex
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHide_Click()
    MessageBox.HideBox
    timUpdate = False
    timUpdate_Timer
    Status = "Status: Unloaded"
End Sub

Private Sub cmdLoad1_Click()
    On Error GoTo Canceled
    CD.FileName = App.Path & "\Blue.bmp"
    CD.ShowOpen
    imgBox.Picture = LoadPicture(CD.FileName)
    txtBox = CD.FileName
    Set MessageBox.BoxBackPic = imgBox.Picture

    Exit Sub
Canceled:
    Exit Sub
End Sub

Private Sub cmdLoad2_Click()
    On Error GoTo Canceled
    CD.FileName = App.Path & "\Green.bmp"
    CD.ShowOpen
    imgTextPic.Picture = LoadPicture(CD.FileName)
    txtSmallPic = CD.FileName
    Set MessageBox.TextBackPic = imgTextPic.Picture

    Exit Sub
Canceled:
    Exit Sub
End Sub

Private Sub cmdShow_Click()
    MessageBox.ShowBox
    timUpdate = True
    timUpdate_Timer
    GetInfo
    Status = "Status: Loaded"
End Sub

Private Sub GetInfo() 'Load information from MsgBox
    'Old part
    txtTitle = MessageBox.TitleCaption
    txtButton1 = MessageBox.Btn1Caption
    txtButton2 = MessageBox.Btn2Caption
    txtButton3 = MessageBox.Btn3Caption
    txtButton4 = MessageBox.Btn4Caption
    txtMessage = MessageBox.MessageText
    cboBStyle.ListIndex = MessageBox.BorderStyle
    
    'Color part
    minBox.BackColor = MessageBox.FormBackColor
    minClose.BackColor = MessageBox.CloseBackColor
    minClose.ForeColor = MessageBox.CloseColor
    minIcon.BackColor = MessageBox.IconBackColor
    minHighLight.BackColor = MessageBox.HighlightButtonColor
    minLowLight.BackColor = MessageBox.NonHighlightButtonColor
    minText.BackColor = MessageBox.MessageTextColor
    minTitle.BackColor = MessageBox.TitleBackColor
    
    'Font part
    minText.Font = MessageBox.MessageFont
    minTitle.Font = MessageBox.TitleFont
End Sub

Private Sub Form_Load()
    GetInfo
End Sub

Private Sub MessageBox_Btn1Click()
    MsgBox "User clicked button 1."
    cmdHide_Click
    Status = "Status: Button 1 Clicked"
End Sub

Private Sub MessageBox_Btn2Click()
    MsgBox "User clicked button 2."
    cmdHide_Click
    Status = "Status: Button 2 Clicked"
End Sub

Private Sub MessageBox_Btn3Click()
    MsgBox "User clicked button 3."
    cmdHide_Click
    Status = "Status: Button 3 Clicked"
End Sub

Private Sub MessageBox_Btn4Click()
    MsgBox "User clicked button 4."
    cmdHide_Click
    Status = "Status: Button 4 Clicked"
End Sub

Private Sub MessageBox_Btn1Focus()
    Status = "Status: Button 1 Recieved Focus"
    Button1.BackColor = MessageBox.HighlightButtonColor
    Button2.BackColor = MessageBox.NonHighlightButtonColor
    Button3.BackColor = MessageBox.NonHighlightButtonColor
    Button4.BackColor = MessageBox.NonHighlightButtonColor
End Sub

Private Sub MessageBox_Btn2Focus()
    Status = "Status: Button 2 Recieved Focus"
    Button2.BackColor = MessageBox.HighlightButtonColor
    Button1.BackColor = MessageBox.NonHighlightButtonColor
    Button3.BackColor = MessageBox.NonHighlightButtonColor
    Button4.BackColor = MessageBox.NonHighlightButtonColor
End Sub

Private Sub MessageBox_Btn3Focus()
    Status = "Status: Button 3 Recieved Focus"
    Button3.BackColor = MessageBox.HighlightButtonColor
    Button2.BackColor = MessageBox.NonHighlightButtonColor
    Button1.BackColor = MessageBox.NonHighlightButtonColor
    Button4.BackColor = MessageBox.NonHighlightButtonColor
End Sub

Private Sub MessageBox_Btn4Focus()
    Status = "Status: Button 4 Recieved Focus"
    Button4.BackColor = MessageBox.HighlightButtonColor
    Button2.BackColor = MessageBox.NonHighlightButtonColor
    Button3.BackColor = MessageBox.NonHighlightButtonColor
    Button1.BackColor = MessageBox.NonHighlightButtonColor
End Sub

Private Sub MessageBox_Terminated()
    MsgBox "Terminated."
    cmdHide_Click
    Status = "Status: Terminated by user"
End Sub

Private Sub timUpdate_Timer()
    txtLeft = MessageBox.BoxLeft
    txtTop = MessageBox.BoxTop
End Sub

Private Sub txtBox_dblClick()
    On Error Resume Next
    txtBox = "(None)"
    imgBox.Picture = LoadPicture(App.Path & "\Blue.bmp")
    Set MessageBox.BoxBackPic = Nothing
End Sub

Private Sub txtSmallPic_dblClick()
    On Error Resume Next
    txtSmallPic = "(None)"
    imgTextPic.Picture = LoadPicture(App.Path & "\Green.bmp")
    Set MessageBox.TextBackPic = Nothing
End Sub

Private Sub txtButton1_Change()
    MessageBox.Btn1Caption = txtButton1
End Sub

Private Sub txtButton2_Change()
    MessageBox.Btn2Caption = txtButton2
End Sub

Private Sub txtButton3_Change()
    MessageBox.Btn3Caption = txtButton3
End Sub

Private Sub txtButton4_Change()
    MessageBox.Btn4Caption = txtButton4
End Sub

Private Sub txtLeft_Change()
    On Error Resume Next
    MessageBox.BoxLeft = CInt(txtLeft)
End Sub

Private Sub txtMessage_Change()
    MessageBox.MessageText = txtMessage
End Sub

Private Sub txtTitle_Change()
    MessageBox.TitleCaption = txtTitle
End Sub

Private Sub txtTop_Change()
    On Error Resume Next
    MessageBox.BoxTop = CInt(txtTop)
End Sub

'Minimap Input Data --------------------------------------------------------------------------------

Private Sub minTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Canceled
    If Button = 1 Then 'Left Click
        CD.Color = MessageBox.TitleBackColor
        CD.ShowColor
        MessageBox.TitleBackColor = CD.Color
    Else 'Right Click
        CD.FontName = MessageBox.TitleFont.Name
        CD.FontSize = MessageBox.TitleFont.Size
        CD.FontBold = MessageBox.TitleFont.Bold
        CD.FontItalic = MessageBox.TitleFont.Italic
        CD.FontStrikethru = MessageBox.TitleFont.Strikethrough
        CD.FontUnderline = MessageBox.TitleFont.Underline
        CD.ShowFont
        
        MessageBox.TitleFont.Name = CD.FontName
        MessageBox.TitleFont.Size = CD.FontSize
        MessageBox.TitleFont.Bold = CD.FontBold
        MessageBox.TitleFont.Italic = CD.FontItalic
        MessageBox.TitleFont.Strikethrough = CD.FontStrikethru
        MessageBox.TitleFont.Underline = CD.FontUnderline
    End If
        
    Exit Sub
Canceled:
    Exit Sub
End Sub

Private Sub minBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Canceled
    If Button = 1 Then 'Left Click
        CD.Color = MessageBox.TitleBackColor
        CD.ShowColor
        MessageBox.FormBackColor = CD.Color
    Else 'Right Click
        CD.Color = MessageBox.BorderColor
        CD.ShowColor
        MessageBox.BorderColor = CD.Color
    End If
        
    Exit Sub
Canceled:
    Exit Sub
End Sub

Private Sub minClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Canceled
    If Button = 1 Then 'Left Click
        CD.Color = MessageBox.CloseColor
        CD.ShowColor
        MessageBox.CloseColor = CD.Color
    Else 'Right Click
        CD.Color = MessageBox.CloseBackColor
        CD.ShowColor
        MessageBox.CloseBackColor = CD.Color
    End If
        
    Exit Sub
Canceled:
    Exit Sub
End Sub

Private Sub minText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Canceled
    If Button = 1 Then 'Left Click
        CD.Color = MessageBox.MessageTextColor
        CD.ShowColor
        MessageBox.MessageTextColor = CD.Color
    Else 'Right Click
        CD.FontName = MessageBox.MessageFont.Name
        CD.FontSize = MessageBox.MessageFont.Size
        CD.FontBold = MessageBox.MessageFont.Bold
        CD.FontItalic = MessageBox.MessageFont.Italic
        CD.FontStrikethru = MessageBox.MessageFont.Strikethrough
        CD.FontUnderline = MessageBox.MessageFont.Underline
        CD.ShowFont
        
        MessageBox.MessageFont.Name = CD.FontName
        MessageBox.MessageFont.Size = CD.FontSize
        MessageBox.MessageFont.Bold = CD.FontBold
        MessageBox.MessageFont.Italic = CD.FontItalic
        MessageBox.MessageFont.Strikethrough = CD.FontStrikethru
        MessageBox.MessageFont.Underline = CD.FontUnderline
    End If
        
    Exit Sub
Canceled:
    Exit Sub
End Sub

Private Sub minIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Canceled
    CD.Color = MessageBox.IconBackColor
    CD.ShowColor
    MessageBox.IconBackColor = CD.Color
        
    Exit Sub
Canceled:
    Exit Sub
End Sub

Private Sub minHighLight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Canceled
    CD.Color = MessageBox.HighlightButtonColor
    CD.ShowColor
    MessageBox.HighlightButtonColor = CD.Color
        
    Exit Sub
Canceled:
    Exit Sub
End Sub

Private Sub minLowLight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Canceled
    CD.Color = MessageBox.NonHighlightButtonColor
    CD.ShowColor
    MessageBox.NonHighlightButtonColor = CD.Color
        
    Exit Sub
Canceled:
    Exit Sub
End Sub
