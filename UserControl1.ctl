VERSION 5.00
Begin VB.UserControl ACPRibbon 
   BackColor       =   &H00404040&
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ScaleHeight     =   3705
   ScaleWidth      =   7095
   Begin VB.Label ButMouse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   4320
      TabIndex        =   9
      ToolTipText     =   "çlll"
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Glip_on 
      Height          =   60
      Index           =   0
      Left            =   4560
      Top             =   2280
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Glip_off 
      Height          =   60
      Index           =   0
      Left            =   4440
      Top             =   2280
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Button_left_over 
      Height          =   990
      Index           =   0
      Left            =   4800
      Top             =   2520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Button_center_over 
      Height          =   990
      Index           =   0
      Left            =   4920
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Button_right_over 
      Height          =   990
      Index           =   0
      Left            =   5760
      Top             =   2520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Cat_Dlg_over 
      Height          =   210
      Index           =   0
      Left            =   4800
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Cat_Dlg_on 
      Height          =   210
      Index           =   0
      Left            =   4560
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Cat_Dlg 
      Height          =   210
      Index           =   0
      Left            =   4320
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Button_Icon 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   0
      Left            =   3600
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Button_Caption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   3720
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image RibbonTopCustom_over 
      Height          =   390
      Left            =   4680
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image RibbonTopCustom 
      Height          =   390
      Left            =   4440
      Top             =   480
      Width           =   225
   End
   Begin VB.Image Button_right 
      Height          =   990
      Index           =   0
      Left            =   4200
      Top             =   2520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Button_center 
      Height          =   990
      Index           =   0
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Button_left 
      Height          =   990
      Index           =   0
      Left            =   3240
      Top             =   2520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label TBMouse 
      Height          =   390
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image RibbonTopImage 
      Height          =   390
      Index           =   0
      Left            =   3360
      Top             =   480
      Width           =   270
   End
   Begin VB.Image RibbonTop_over 
      Height          =   390
      Index           =   0
      Left            =   3720
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label TabMouse 
      Height          =   360
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Tab_caption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aba 01"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2820
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image Tab_right 
      Height          =   360
      Index           =   0
      Left            =   1560
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_center 
      Height          =   360
      Index           =   0
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Tab_left 
      Height          =   360
      Index           =   0
      Left            =   960
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_left_over 
      Height          =   360
      Index           =   0
      Left            =   960
      Top             =   3240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_center_over 
      Height          =   360
      Index           =   0
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Tab_right_over 
      Height          =   360
      Index           =   0
      Left            =   1560
      Top             =   3240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label CatMouse 
      Height          =   1350
      Index           =   0
      Left            =   5280
      TabIndex        =   7
      Top             =   750
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Cat_Caption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   5760
      TabIndex        =   6
      Tag             =   "sadf"
      Top             =   1800
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image Cat_Right_on 
      Height          =   1335
      Index           =   0
      Left            =   6840
      Top             =   750
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Cat_Center_on 
      Height          =   1335
      Index           =   0
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   750
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Left_on 
      Height          =   1335
      Index           =   0
      Left            =   6480
      Top             =   750
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Right_off 
      Height          =   1335
      Index           =   0
      Left            =   6120
      Top             =   750
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Cat_Left_off 
      Height          =   1335
      Index           =   0
      Left            =   5760
      Top             =   750
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Center_off 
      Height          =   1335
      Index           =   0
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   750
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label ButtonRibbon 
      Height          =   675
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   690
   End
   Begin VB.Image Endon 
      Height          =   345
      Left            =   6240
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Maxon 
      Height          =   345
      Left            =   5520
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Minon 
      Height          =   345
      Left            =   4800
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Endoff 
      Height          =   345
      Left            =   3960
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Maxoff 
      Height          =   345
      Left            =   3240
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Minoff 
      Height          =   345
      Left            =   2520
      Top             =   0
      Width           =   600
   End
   Begin VB.Label Barra 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Titulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   630
   End
   Begin VB.Label Titulo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFD18A&
      Height          =   240
      Left            =   840
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image RibbonTopRight 
      Height          =   390
      Left            =   3120
      Top             =   480
      Width           =   195
   End
   Begin VB.Image RibbonTop 
      Height          =   390
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Logo 
      Height          =   360
      Left            =   2760
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image ButtonRibbonon 
      Height          =   675
      Left            =   1800
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image ButtonRibbonover 
      Height          =   675
      Left            =   1800
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image ButtonRibbonoff 
      Height          =   675
      Left            =   1800
      Top             =   480
      Width           =   735
   End
   Begin VB.Image BarraLeft 
      Height          =   2130
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Image BarraRight 
      Height          =   2130
      Left            =   960
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Barra2 
      Height          =   2130
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "ACPRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'#######################################
'#                                     #
'#           ACP Ribbon 2007           #
'#                  by                 #
'#      adrianopaladini@gmail.com      #
'#                                     #
'#                                     #
'#  Visual from Office 2007 Beta 2 TR  #
'#                                     #
'#   Please Don´t Remove Author Info!  #
'#                                     #
'#######################################


'------------------------------------------------
' TO DO:
'
' A) Update when Resize, resolve flicks
' B) Optimize Code
' C) Insert Mini Buttons, Combos and Checkbox on Each Categories
' D) Option to Show Menu Under the Ribbon and Hide Ribbon
' E) Make Menu
' F) Option to user customize the menu
' G) Group Tabs
' H) Add Comment to All code
' I) FINISHED this project!
'
'------------------------------------------------

'------------------------------------------------
' Bugs:
'
' Please report to:
'
'         adrianopaladini@gmail.com
'
'------------------------------------------------


Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Dim TotalTopButton As Integer
Dim TotalButton As Integer
Dim TotalTabs As Integer
Dim TotalCats As Integer
Dim TabSelected As String
Dim TabID(30) As String
Dim TabC(30) As String
Dim CatsID(30) As String
Dim CatsC(30) As String
Dim CatsT(30) As String
Dim CatsD(30) As Boolean
Dim TopBID(30) As String
Dim TopBC(30) As String

Dim TopBuID(90) As String
Dim TopBuS(90) As String
Dim TopBuC(90) As String
Dim TopBuI(90) As Picture
Dim TopBuT(90) As String
Dim TopBuG(90) As Boolean

Dim MS As Boolean
Dim Mx, My As Integer
Dim sCaption As String
Const m_def_Caption = ""
Const m_def_ShowCustomMenu = False
Dim m_ShowCustomMenu As Boolean
Event MainMenuClick()
Event MenuClick(ByVal ID As String, ByVal Caption As String)
Event TabClick(ByVal ID As String, ByVal Caption As String)
Event CatClick(ByVal ID As String, ByVal Caption As String)
Event ButtonClick(ByVal ID As String, ByVal Caption As String)
Event CustomClick()
Const m_def_Theme = 0
Dim m_Theme As Variant
Dim zImg As ImageList

Dim TAB_NORMAL
Dim TAB_SELECTED
Private Sub Barra_DblClick()
Maxon_Click
End Sub
Private Sub Barra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mx = X
My = Y
MS = True
End Sub
Private Sub Barra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If MS = True Then
    UserControl.ParentControls.Item(0).Move UserControl.ParentControls.Item(0).Left - (Mx - X), UserControl.ParentControls.Item(0).Top - (My - Y)
End If
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub Barra_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MS = False
End Sub
Private Sub Barra2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub BarraLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub BarraRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub

Private Sub ButMouse_Click(Index As Integer)
RaiseEvent ButtonClick(ButMouse(Index).Tag, Button_Caption(Index).Caption)
End Sub

Private Sub ButMouse_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Button_left_over(Index).Visible = True
    Button_center_over(Index).Visible = True
    Button_right_over(Index).Visible = True
End Sub

Private Sub ButMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To ButMouse.UBound
    If i <> Index Then
        Button_left(i).Visible = False
        Button_center(i).Visible = False
        Button_right(i).Visible = False
        If Glip_off(i).Visible = True Then
            Glip_on(i).Visible = False
        End If
    End If
Next
If Button_left(Index).Visible = False Then
    Button_left(Index).Visible = True
    Button_center(Index).Visible = True
    Button_right(Index).Visible = True
    If Glip_off(Index).Visible = True Then
        Glip_on(Index).Visible = True
    End If
End If
For i = 0 To CatMouse.UBound
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_over(i).Visible = False
    End If
Next
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub

Private Sub ButMouse_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Button_left_over(Index).Visible = False
    Button_center_over(Index).Visible = False
    Button_right_over(Index).Visible = False
End Sub

Private Sub ButtonRibbon_Click()
RaiseEvent MainMenuClick
End Sub
Private Sub ButtonRibbon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = True
End Sub
Private Sub ButtonRibbon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonRibbonover.Visible = True
ButtonRibbonon.Visible = False
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
End Sub
Private Sub ButtonRibbon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonRibbonover.Visible = True
ButtonRibbonon.Visible = False
End Sub


Private Sub Cat_Dlg_on_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Cat_Dlg_over(Index).Visible = True
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

End Sub

Private Sub Cat_Dlg_over_Click(Index As Integer)
    RaiseEvent CatClick(Cat_Caption(Index).Tag, Cat_Caption(Index).Caption)
End Sub

Private Sub CatMouse_Click(Index As Integer)
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
Next

    Cat_Center_on(Index).Visible = True
    Cat_Left_on(Index).Visible = True
    Cat_Right_on(Index).Visible = True
End Sub
Private Sub CatMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To CatMouse.UBound
    If i = Index Then
        If Cat_Center_on(i).Visible = False Then
            Cat_Center_on(Index).Visible = True
            Cat_Left_on(Index).Visible = True
            Cat_Right_on(Index).Visible = True
            If Cat_Dlg(i).Visible = True Then
                Cat_Dlg_on(Index).Visible = True
            End If
        End If
    Else
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    End If
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub Endon_Click()
Endon.Visible = False
Unload UserControl.ParentControls.Item(0)
End Sub

Private Sub Maxon_Click()
If UserControl.ParentControls.Item(0).WindowState = 2 Then
    UserControl.ParentControls.Item(0).WindowState = 0
Else
    UserControl.ParentControls.Item(0).WindowState = 2
End If
Maxon.Visible = False
UserControl_Resize
End Sub
Private Sub Minoff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = True
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub Maxoff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Endon.Visible = False
Maxon.Visible = True
Minon.Visible = False
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub Endoff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Endon.Visible = True
Maxon.Visible = False
Minon.Visible = False
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub Minon_Click()
UserControl.ParentControls.Item(0).WindowState = 1
Minon.Visible = False
End Sub

Private Sub RibbonTopCustom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RibbonTopCustom_over.Visible = True
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
End Sub
Private Sub RibbonTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RibbonTopCustom_over.Visible = False
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next

For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
End Sub
Private Sub TabMouse_Click(Index As Integer)
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
    
    Tab_center(i).Visible = False
    Tab_left(i).Visible = False
    Tab_right(i).Visible = False
    
    Tab_caption(i).ForeColor = TAB_NORMAL
Next
Tab_caption(Index).ForeColor = TAB_SELECTED
Tab_center(Index).Visible = True
Tab_left(Index).Visible = True
Tab_right(Index).Visible = True
TabSelected = TabID(Index)
CatsUpdate
RaiseEvent TabClick(TabID(Index), TabC(Index))
End Sub
Private Sub TabMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To TabMouse.UBound
    If i = Index Then
        If Tab_center(i).Visible = False Then
            Tab_center_over(Index).Visible = True
            Tab_left_over(Index).Visible = True
            Tab_right_over(Index).Visible = True
        End If
    Else
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    End If
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next
For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub TBMouse_Click(Index As Integer)
    RaiseEvent MenuClick(TopBID(Index), TopBC(Index))
End Sub
Private Sub TBMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To TBMouse.UBound
    RibbonTop_over(i).Visible = False
Next
RibbonTop_over(Index).Visible = True
For i = 0 To TabMouse.UBound
    Tab_center_over(i).Visible = False
    Tab_left_over(i).Visible = False
    Tab_right_over(i).Visible = False
Next
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_right(KL).Visible = False
    Button_center(KL).Visible = False
Next
For i = 0 To CatMouse.UBound
    Cat_Center_on(i).Visible = False
    Cat_Left_on(i).Visible = False
    Cat_Right_on(i).Visible = False
    If Cat_Dlg(i).Visible = True Then
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
    End If
Next
RibbonTopCustom_over.Visible = False
Endon.Visible = False
Maxon.Visible = False
Minon.Visible = False
ButtonRibbonover.Visible = False
ButtonRibbonon.Visible = False
End Sub
Private Sub UserControl_Initialize()
TotalTopButton = 0
TotalButton = 0
TotalTabs = 0
TotalCats = 0

TabSelected = ""

Barra.BackStyle = 0
ButtonRibbon.BackStyle = 0
TabMouse(0).BackStyle = 0
CatMouse(0).BackStyle = 0
TBMouse(0).BackStyle = 0
ButMouse(0).BackStyle = 0

UserControl_Resize
TabsUpdate
CatsUpdate




End Sub
Private Sub TabsUpdate()
On Error Resume Next
For i = 1 To (TotalTabs - 1)
        Unload Tab_caption(i)
        Unload Tab_left(i)
        Unload Tab_center(i)
        Unload Tab_right(i)
        Unload Tab_left_over(i)
        Unload Tab_center_over(i)
        Unload Tab_right_over(i)
        Unload TabMouse(i)
Next
For i = 0 To (TotalTabs - 1)
    If i <> 0 Then
        Load Tab_caption(i)
        Load Tab_left(i)
        Load Tab_center(i)
        Load Tab_right(i)
        Load Tab_left_over(i)
        Load Tab_center_over(i)
        Load Tab_right_over(i)
        Load TabMouse(i)
        Tab_left(i).Left = Tab_right(i - 1).Left + Tab_right(i).Width
    Else
        Tab_left(0).Left = ButtonRibbon.Width
    End If
    TabMouse(i).Left = Tab_left(i).Left
    
    Tab_caption(i).Top = 395 + 60
    Tab_center(i).Top = 395
    Tab_left(i).Top = 395
    Tab_right(i).Top = 395
    Tab_center_over(i).Top = 395
    Tab_left_over(i).Top = 395
    Tab_right_over(i).Top = 395
    TabMouse(i).Top = 395
    
    Tab_caption(i) = TabC(i)
    Tab_center(i).Width = Tab_caption(i).Width
    Tab_center(i).Left = Tab_left(i).Left + Tab_left(i).Width
    Tab_caption(i).Left = Tab_center(i).Left
    Tab_right(i).Left = Tab_center(i).Left + Tab_center(i).Width
    
    Tab_center_over(i).Width = Tab_center(i).Width
    Tab_center_over(i).Left = Tab_center(i).Left
    Tab_left_over(i).Left = Tab_left(i).Left
    Tab_right_over(i).Left = Tab_right(i).Left
    
    TabMouse(i).Width = Tab_left(i).Width + Tab_right(i).Width + Tab_center(i).Width
    
    Tab_caption(i).ForeColor = TAB_NORMAL
    
    Tab_caption(i).Visible = True
    If i = 0 Then
        Tab_center(i).Visible = True
        Tab_left(i).Visible = True
        Tab_right(i).Visible = True
        Tab_caption(i).ForeColor = TAB_SELECTED
    End If
    TabMouse(i).Visible = True

    Tab_center(i).ZOrder 0
    Tab_left(i).ZOrder 0
    Tab_right(i).ZOrder 0
    
    Tab_center_over(i).ZOrder 0
    Tab_left_over(i).ZOrder 0
    Tab_right_over(i).ZOrder 0
    
    Tab_caption(i).ZOrder 0
    TabMouse(i).ZOrder 0
Next
End Sub
Private Sub CatsUpdate()
'On Error Resume Next
Dim TotalCatsT As Integer
Dim CatsIDT(30) As String
Dim CatsCT(30) As String
Dim CatsTT(30) As String
Dim CatsDT(30) As Boolean
TotalCatsT = 0
For i = 0 To TotalCats
    If CatsT(i) = TabSelected And TabSelected <> "" And CatsT(i) <> "" Then
        CatsIDT(TotalCatsT) = CatsID(i)
        CatsTT(TotalCatsT) = CatsT(i)
        CatsCT(TotalCatsT) = CatsC(i)
        CatsDT(TotalCatsT) = CatsD(i)
        TotalCatsT = TotalCatsT + 1
    End If
Next
For i = 1 To CatMouse.UBound
        Unload Cat_Left_off(i)
        Unload Cat_Left_on(i)
        Unload Cat_Right_off(i)
        Unload Cat_Right_on(i)
        Unload Cat_Center_off(i)
        Unload Cat_Center_on(i)
        Unload Cat_Caption(i)
        Unload CatMouse(i)
        Unload Cat_Dlg(i)
        Unload Cat_Dlg_on(i)
        Unload Cat_Dlg_over(i)
Next
For i = 1 To Button_center.UBound
    Unload Button_left(i)
    Unload Button_center(i)
    Unload Button_right(i)
    Unload Button_left_over(i)
    Unload Button_center_over(i)
    Unload Button_right_over(i)
    Unload Button_Caption(i)
    Unload Button_Icon(i)
    Unload Glip_on(i)
    Unload Glip_off(i)
    Unload ButMouse(i)
Next
Button_left(0).Visible = False
Button_center(0).Visible = False
Button_right(0).Visible = False
Button_Caption(0).Visible = False
Button_Icon(0).Visible = False
ButMouse(0).Visible = False

Cat_Left_off(0).Visible = False
Cat_Left_on(0).Visible = False
Cat_Right_off(0).Visible = False
Cat_Right_on(0).Visible = False
Cat_Center_off(0).Visible = False
Cat_Center_on(0).Visible = False
Cat_Caption(0).Visible = False
CatMouse(0).Visible = False
Cat_Dlg(0).Visible = False
Cat_Dlg_on(0).Visible = False
Cat_Dlg_over(0).Visible = False
For i = 0 To (TotalCatsT - 1)
    If i <> 0 Then
        Load Cat_Left_off(i)
        Load Cat_Left_on(i)
        Load Cat_Right_off(i)
        Load Cat_Right_on(i)
        Load Cat_Center_off(i)
        Load Cat_Center_on(i)
        Load Cat_Caption(i)
        Load CatMouse(i)
        Load Cat_Dlg(i)
        Load Cat_Dlg_on(i)
        Load Cat_Dlg_over(i)
        Cat_Left_off(i).Left = Cat_Right_off(i - 1).Left + Cat_Right_off(i).Width
    Else
        Cat_Left_off(i).Left = 120
    End If
    CatMouse(i).Left = Cat_Left_off(i).Left
    
    Cat_Caption(i).Caption = CatsCT(i)
    Cat_Caption(i).Tag = CatsIDT(i)
    
    Cat_Center_off(i).Left = Cat_Left_off(i).Left + Cat_Left_off(i).Width
    
    BUTSIZE = ButtonsUpdate(CatsIDT(i), Cat_Center_off(i).Left)
    
    If CatsDT(i) = True Then
        Cat_Center_off(i).Width = Cat_Caption(i).Width + Cat_Dlg(i).Width
    Else
        Cat_Center_off(i).Width = Cat_Caption(i).Width
    End If
    
    If Cat_Center_off(i).Width < BUTSIZE Then
        Cat_Center_off(i).Width = BUTSIZE
        Cat_Caption(i).Left = Cat_Center_off(i).Left + ((Cat_Center_off(i).Width - Cat_Caption(i).Width) / 2)
    Else
        Cat_Caption(i).Left = Cat_Center_off(i).Left
    End If
    
    Cat_Right_off(i).Left = Cat_Center_off(i).Left + Cat_Center_off(i).Width
    
    Cat_Center_on(i).Width = Cat_Center_off(i).Width
    Cat_Center_on(i).Left = Cat_Center_off(i).Left
    Cat_Left_on(i).Left = Cat_Left_off(i).Left
    Cat_Right_on(i).Left = Cat_Right_off(i).Left
    
    CatMouse(i).Width = Cat_Left_off(i).Width + Cat_Right_off(i).Width + Cat_Center_off(i).Width
    
    Cat_Caption(i).Visible = True
    Cat_Center_off(i).Visible = True
    Cat_Left_off(i).Visible = True
    Cat_Right_off(i).Visible = True
    CatMouse(i).Visible = True

    Cat_Center_off(i).ZOrder 0
    Cat_Left_off(i).ZOrder 0
    Cat_Right_off(i).ZOrder 0
    
    Cat_Center_on(i).ZOrder 0
    Cat_Left_on(i).ZOrder 0
    Cat_Right_on(i).ZOrder 0
    
    Cat_Caption(i).ZOrder 0
    CatMouse(i).ZOrder 0
    
    Cat_Dlg(i).Left = (Cat_Right_off(i).Left - Cat_Dlg(i).Width) + 15
    Cat_Dlg(i).Top = (Cat_Right_off(i).Top + Cat_Right_off(i).Height) - (Cat_Dlg(i).Height + 60)
    
    Cat_Dlg_on(i).Left = Cat_Dlg(i).Left
    Cat_Dlg_over(i).Left = Cat_Dlg(i).Left
    
    Cat_Dlg_on(i).Top = Cat_Dlg(i).Top
    Cat_Dlg_over(i).Top = Cat_Dlg(i).Top
    
    
    Cat_Dlg_on(i).Visible = False
    Cat_Dlg_over(i).Visible = False
    
    If CatsDT(i) = True Then
        Cat_Dlg(i).Visible = True
    End If
    Cat_Dlg(i).ZOrder 0
    Cat_Dlg_on(i).ZOrder 0
    Cat_Dlg_over(i).ZOrder 0
Next
DoEvents
For KL = 0 To ButMouse.UBound
    Button_left(KL).Visible = False
    Button_left(KL).ZOrder 0
    Button_right(KL).Visible = False
    Button_right(KL).ZOrder 0
    Button_center(KL).Visible = False
    Button_center(KL).ZOrder 0
    
    Button_left_over(KL).Visible = False
    Button_left_over(KL).ZOrder 0
    Button_right_over(KL).Visible = False
    Button_right_over(KL).ZOrder 0
    Button_center_over(KL).Visible = False
    Button_center_over(KL).ZOrder 0
    
    Button_Icon(KL).ZOrder 0
    Button_Caption(KL).ZOrder 0
    
    Glip_off(KL).ZOrder 0
    Glip_on(KL).ZOrder 0
    
    ButMouse(KL).ZOrder 0
Next

End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.Height = Barra2.Height
    UserControl.Width = UserControl.ParentControls.Item(0).ScaleWidth
    Barra2.Width = UserControl.Width
    BarraRight.Left = Barra2.Width - BarraRight.Width
    
    ButtonRibbon.Top = 0
    ButtonRibbon.Left = 0
    
    ButtonRibbonoff.Top = 0
    ButtonRibbonover.Top = 0
    ButtonRibbonon.Top = 0
    
    ButtonRibbonoff.Left = 0
    ButtonRibbonover.Left = 0
    ButtonRibbonon.Left = 0
    
    Logo.Top = (ButtonRibbonoff.Height - Logo.Height) / 2
    Logo.Left = Logo.Top
    
    
    RibbonTop.Top = 0
    RibbonTop.Left = ButtonRibbonoff.Width
    RibbonTopImage(TotalTopButton - 1).Top = (RibbonTop.Height - RibbonTopImage(TotalTopButton - 1).Height) / 2
    
    RibbonTop.Width = 330 * TotalTopButton
    
    RibbonTopRight.Top = 0
    RibbonTopRight.Left = RibbonTop.Left + RibbonTop.Width
    
    RibbonTopCustom.Top = 0
    RibbonTopCustom.Left = RibbonTopRight.Left + RibbonTopRight.Width
    RibbonTopCustom_over.Top = 0
    RibbonTopCustom_over.Left = RibbonTopCustom.Left
    
    If m_ShowCustomMenu = True Then
        RibbonTopCustom.Visible = True
        InicioArea = (RibbonTopCustom.Left + RibbonTopCustom.Width)
    Else
        RibbonTopCustom.Visible = False
        InicioArea = (RibbonTopRight.Left + RibbonTopRight.Width)
    End If
    
    
    area = UserControl.Width - (InicioArea + (Endoff.Width * 3))
    
    Barra.Left = InicioArea
    Barra.Width = area

    Pos = InStr(sCaption, " - ")
    If Pos > 0 Then
        Titulo.Caption = Mid(sCaption, 1, Pos + 2)
        Titulo2.Caption = Mid(sCaption, Pos + 3)
        Titulo.Left = ((area - (Titulo.Width + Titulo2.Width)) / 2) + InicioArea
        Titulo2.Left = Titulo.Left + Titulo.Width
        Titulo2.Visible = True
    Else
        Titulo.Caption = sCaption
        Titulo.Left = ((area - Titulo.Width) / 2) + InicioArea
        Titulo2.Visible = False
    End If
    
    Endoff.Left = Barra2.Width - Endoff.Width
    Endon.Left = Endoff.Left
    Maxoff.Left = Endoff.Left - Maxoff.Width
    Maxon.Left = Maxoff.Left
    Minoff.Left = Maxoff.Left - Minoff.Width
    Minon.Left = Minoff.Left
End Sub
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Caption = sCaption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    sCaption = New_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     UserControl_Resize
     TabsUpdate
     CatsUpdate
End Sub
Private Sub UserControl_InitProperties()
    sCaption = m_def_Caption
    m_ShowCustomMenu = m_def_ShowCustomMenu
    m_Theme = m_def_Theme
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    sCaption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_ShowCustomMenu = PropBag.ReadProperty("ShowCustomMenu", m_def_ShowCustomMenu)
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", sCaption, m_def_Caption)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ShowCustomMenu", m_ShowCustomMenu, m_def_ShowCustomMenu)
    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
End Sub
Public Function AddTab(zID As String, zCaption As String) As Boolean
    TotalTabs = TotalTabs + 1
    TabID(TotalTabs - 1) = zID
    zCaption = Replace(zCaption, vbNewLine, " ")
    TabC(TotalTabs - 1) = zCaption
    If TabSelected = "" Then
        TabSelected = zID
    End If
End Function
Public Function AddCat(zID As String, zTab As String, zCaption As String, zDlgButton As Boolean) As Boolean
    TotalCats = TotalCats + 1
    CatsID(TotalCats - 1) = zID
    CatsT(TotalCats - 1) = zTab
    zCaption = Replace(zCaption, vbNewLine, " ")
    CatsC(TotalCats - 1) = zCaption
    CatsD(TotalCats - 1) = zDlgButton
End Function
Public Function AddTopButton(zID As String, zCaption As String, zPicture As Integer) As Boolean
    TotalTopButton = TotalTopButton + 1
    TopBID(TotalTopButton - 1) = zID
    TopBC(TotalTopButton - 1) = zCaption
    If TotalTopButton <> 1 Then
        Load RibbonTopImage(TotalTopButton - 1)
        Load RibbonTop_over(TotalTopButton - 1)
        Load TBMouse(TotalTopButton - 1)
    End If
    TBMouse(TotalTopButton - 1).Top = 0
    RibbonTop_over(TotalTopButton - 1).Top = 0
    RibbonTop_over(TotalTopButton - 1).Left = RibbonTop.Left + (330 * (TotalTopButton - 1))
    TBMouse(TotalTopButton - 1).Left = RibbonTop_over(TotalTopButton - 1).Left
    Set RibbonTopImage(TotalTopButton - 1) = zImg.ListImages.Item(zPicture).Picture
    RibbonTopImage(TotalTopButton - 1).Top = (RibbonTop.Height - RibbonTopImage(TotalTopButton - 1).Height) / 2
    
    ct = (RibbonTop_over(TotalTopButton - 1).Width - RibbonTopImage(TotalTopButton - 1).Width) / 2
    RibbonTopImage(TotalTopButton - 1).Left = RibbonTop_over(TotalTopButton - 1).Left + ct
    
    RibbonTop_over(TotalTopButton - 1).Visible = False
    RibbonTop_over(TotalTopButton - 1).ZOrder 0
    RibbonTopImage(TotalTopButton - 1).Visible = True
    RibbonTopImage(TotalTopButton - 1).ZOrder 0
    zCaption = Replace(zCaption, vbNewLine, " ")
    TBMouse(TotalTopButton - 1).ToolTipText = zCaption
    TBMouse(TotalTopButton - 1).Visible = True
    TBMouse(TotalTopButton - 1).ZOrder 0
End Function
Public Property Get ShowCustomMenu() As Boolean
Attribute ShowCustomMenu.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    ShowCustomMenu = m_ShowCustomMenu
End Property
Public Property Let ShowCustomMenu(ByVal New_ShowCustomMenu As Boolean)
    m_ShowCustomMenu = New_ShowCustomMenu
    PropertyChanged "ShowCustomMenu"
End Property
Private Sub RibbonTopCustom_over_Click()
    RaiseEvent CustomClick
End Sub

Public Function AddButton(zID As String, zSubCat As String, zCaption As String, zPicture As Integer, Optional zMore As Boolean = False, Optional zToolTip As String) As Boolean
    TotalButton = TotalButton + 1
    TopBuID(TotalButton - 1) = zID
    TopBuS(TotalButton - 1) = zSubCat
    TopBuC(TotalButton - 1) = zCaption
    If zToolTip = "" Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace(zCaption, vbNewLine, " ")
        End If
        TopBuT(TotalButton - 1) = zCaption
    Else
        zToolTip = Replace(zToolTip, vbNewLine, " ")
        TopBuT(TotalButton - 1) = zToolTip
    End If
    Set TopBuI(TotalButton - 1) = zImg.ListImages.Item(zPicture).Picture
    TopBuG(TotalButton - 1) = zMore
End Function

Private Function ButtonsUpdate(SubCat As String, PosIni As Integer) As Integer
'On Error Resume Next
Dim TotalButtonT As Integer
Dim TopBuIDT(90) As String
Dim TopBuST(90) As String
Dim TopBuCT(90) As String
Dim TopBuIT(90) As Picture
Dim TopBuTT(90) As String
Dim TopBuGT(90) As Boolean
TotalSize = 0
TotalButtonT = 0
For i = 0 To TotalButton
    If TopBuS(i) = SubCat Then
        TopBuIDT(TotalButtonT) = TopBuID(i)
        TopBuST(TotalButtonT) = TopBuS(i)
        TopBuCT(TotalButtonT) = TopBuC(i)
        TopBuTT(TotalButtonT) = TopBuT(i)
        Set TopBuIT(TotalButtonT) = TopBuI(i)
        TopBuGT(TotalButtonT) = TopBuG(i)
        TotalButtonT = TotalButtonT + 1
    End If
Next
Button_left(0).Visible = False
Button_center(0).Visible = False
Button_right(0).Visible = False
Button_Caption(0).Visible = True
Button_Icon(0).Visible = True
ButMouse(0).Visible = True

xt = ButMouse.UBound + 1

For i = xt To (TotalButtonT - 1) + xt
    If i <> 0 Then
        Load Button_left(i)
        Load Button_center(i)
        Load Button_right(i)
        Load Button_left_over(i)
        Load Button_center_over(i)
        Load Button_right_over(i)
        Load Button_Caption(i)
        Load Button_Icon(i)
        Load Glip_on(i)
        Load Glip_off(i)
        Load ButMouse(i)
    End If
    ButMouse(i).Tag = TopBuIDT(i - xt)
    ButMouse(i).Top = Cat_Left_off(0).Top + 60
    Button_left(i).Top = ButMouse(i).Top
    Button_center(i).Top = ButMouse(i).Top
    Button_right(i).Top = ButMouse(i).Top
    Button_left_over(i).Top = ButMouse(i).Top
    Button_center_over(i).Top = ButMouse(i).Top
    Button_right_over(i).Top = ButMouse(i).Top
    
    If i = xt Then
        posatu = PosIni
    Else
        posatu = ButMouse(i - 1).Left + ButMouse(i - 1).Width + 30
    End If
    ButMouse(i).Left = posatu
    Button_left(i).Left = ButMouse(i).Left
    Button_left_over(i).Left = Button_left(i).Left
    Button_center(i).Left = Button_left(i).Left + Button_left(i).Width
    Button_center_over(i).Left = Button_center(i).Left
    
    Button_Caption(i).Caption = TopBuCT(i - xt)
    
    Set Button_Icon(i) = TopBuIT(i - xt)
    
    ESP = Button_center(i).Height - (Button_Icon(i).Height + Button_Caption(i).Height)
    
    
    
    If TopBuGT(i - xt) = True Then
        Button_Icon(i).Top = Button_center(i).Top + ((ESP - (Button_Caption(i).Height / 2)) / 2)
    Else
        Button_Icon(i).Top = Button_center(i).Top + ((ESP) / 2)
    End If
    Button_Caption(i).Top = Button_Icon(i).Top + Button_Icon(i).Height
    
    Glip_off(i).Top = Button_Caption(i).Top + Button_Caption(i).Height + ((Button_Caption(i).Height - Glip_off(i).Height) / 2)
    Glip_on(i).Top = Glip_off(i).Top
    
    
    If Button_Caption(i).Width > Button_Icon(i).Width Then
        Button_Caption(i).Left = Button_center(i).Left
        esp2 = (Button_Caption(i).Width - Button_Icon(i).Width) / 2
        Button_Icon(i).Left = Button_Caption(i).Left + esp2
        area = Button_Caption(i).Width
    Else
        Button_Icon(i).Left = Button_center(i).Left
        esp2 = (Button_Icon(i).Width - Button_Caption(i).Width) / 2
        Button_Caption(i).Left = Button_Icon(i).Left + esp2
        area = Button_Icon(i).Width
    End If

    Glip_off(i).Left = Button_Caption(i).Left + ((Button_Caption(i).Width - Glip_on(i).Width) / 2)
    Glip_on(i).Left = Glip_off(i).Left

    Button_center(i).Width = area
    Button_center_over(i).Width = Button_center(i).Width
    Button_right(i).Left = Button_center(i).Left + Button_center(i).Width
    Button_right_over(i).Left = Button_right(i).Left
    ButMouse(i).Width = (Button_right(i).Width + Button_right(i).Width) + Button_center(i).Width
    
    ButMouse(i).ToolTipText = TopBuTT(i - xt)
    Button_Icon(i).Visible = True
    Button_Caption(i).Visible = True
    ButMouse(i).Visible = True
    If TopBuGT(i - xt) = True Then
        Glip_off(i).Visible = True
        Glip_off(i).ZOrder 0
        Glip_on(i).ZOrder 0
    End If
    
    TotalSize = TotalSize + ButMouse(i).Width + 30
Next
ButtonsUpdate = TotalSize - 30

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Theme() As Variant
Attribute Theme.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As Variant)
    m_Theme = New_Theme
    PropertyChanged "Theme"
    LoadTheme m_Theme
End Property

Function LoadTheme(iTema)

Select Case iTema
    Case 0
        ID = "BLACK"
        Titulo.ForeColor = &HFFFFFF
        Titulo2.ForeColor = &HFFD18A
        Cat_Caption(0).ForeColor = &HFFFFFF
        TAB_NORMAL = vbWhite
        TAB_SELECTED = vbBlack
        Button_Caption(0).ForeColor = &H80000008
    Case 1
        ID = "BLUE"
        Titulo.ForeColor = &H797069
        Titulo2.ForeColor = &HB86A3E
        Cat_Caption(0).ForeColor = &HB86A3E
        TAB_NORMAL = &H8B4215
        TAB_SELECTED = &H8B4215
        Button_Caption(0).ForeColor = &H8B4215
    Case 2
        ID = "SILVER"
        Titulo.ForeColor = &H6A625C
        Titulo2.ForeColor = &HB86A3E
        Cat_Caption(0).ForeColor = &H6A625C
        TAB_NORMAL = &H6A625C
        TAB_SELECTED = &H6A625C
        Button_Caption(0).ForeColor = &H6A625C
    Case Else
        ID = "BLACK"
End Select

   Set Barra2.Picture = LoadResPicture(101, ID)
   Set BarraLeft.Picture = LoadResPicture(102, ID)
   Set BarraRight.Picture = LoadResPicture(103, ID)
   Set Minoff.Picture = LoadResPicture(104, ID)
   Set Minon.Picture = LoadResPicture(105, ID)
   Set Maxoff.Picture = LoadResPicture(106, ID)
   Set Maxon.Picture = LoadResPicture(107, ID)
   Set Endoff.Picture = LoadResPicture(108, ID)
   Set Endon.Picture = LoadResPicture(109, ID)
   Set ButtonRibbonoff.Picture = LoadResPicture(110, ID)
   Set ButtonRibbonover.Picture = LoadResPicture(111, ID)
   Set ButtonRibbonon.Picture = LoadResPicture(112, ID)
   Set RibbonTop.Picture = LoadResPicture(113, ID)
   Set RibbonTopRight.Picture = LoadResPicture(114, ID)
   Set RibbonTopCustom.Picture = LoadResPicture(115, ID)
   Set RibbonTopCustom_over.Picture = LoadResPicture(116, ID)

   Set RibbonTop_over(0).Picture = LoadResPicture(117, ID)
   Set Cat_Dlg(0).Picture = LoadResPicture(118, ID)
   Set Cat_Dlg_on(0).Picture = LoadResPicture(119, ID)
   Set Cat_Dlg_over(0).Picture = LoadResPicture(120, ID)
   Set Cat_Left_off(0).Picture = LoadResPicture(121, ID)
   Set Cat_Center_off(0).Picture = LoadResPicture(122, ID)
   Set Cat_Right_off(0).Picture = LoadResPicture(123, ID)
   Set Cat_Left_on(0).Picture = LoadResPicture(124, ID)
   Set Cat_Center_on(0).Picture = LoadResPicture(125, ID)
   Set Cat_Right_on(0).Picture = LoadResPicture(126, ID)
   Set Tab_left(0).Picture = LoadResPicture(127, ID)
   Set Tab_center(0).Picture = LoadResPicture(128, ID)
   Set Tab_right(0).Picture = LoadResPicture(129, ID)
   Set Tab_left_over(0).Picture = LoadResPicture(130, ID)
   Set Tab_center_over(0).Picture = LoadResPicture(131, ID)
   Set Tab_right_over(0).Picture = LoadResPicture(132, ID)
   Set Glip_off(0).Picture = LoadResPicture(133, ID)
   Set Glip_on(0).Picture = LoadResPicture(134, ID)
   Set Button_left_over(0).Picture = LoadResPicture(135, ID)
   Set Button_center_over(0).Picture = LoadResPicture(136, ID)
   Set Button_right_over(0).Picture = LoadResPicture(137, ID)
   Set Button_left(0).Picture = LoadResPicture(138, ID)
   Set Button_center(0).Picture = LoadResPicture(139, ID)
   Set Button_right(0).Picture = LoadResPicture(140, ID)
    
End Function

Private Property Get TempDir() As String
Dim sRet As String, c As Long
Dim lErr As Long
   sRet = String$(MAX_PATH, 0)
   c = GetTempPath(MAX_PATH, sRet)
   lErr = Err.LastDllError
   If c = 0 Then
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
   TempDir = Left$(sRet, c)
End Property

Private Property Get TempFileName( _
        Optional ByVal sPrefix As String, _
        Optional ByVal sPathName As String) As String
Dim lErr As Long
Dim iPos As Long

   If sPrefix = "" Then sPrefix = ""
   If sPathName = "" Then sPathName = TempDir
   
   Dim sRet As String
   sRet = String(MAX_PATH, 0)
   GetTempFileName sPathName, sPrefix, 0, sRet
   lErr = Err.LastDllError
   If Not lErr = 0 Then
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
   iPos = InStr(sRet, vbNullChar)
   If Not iPos = 0 Then
      TempFileName = Left$(sRet, iPos - 1)
   End If
End Property

Private Function WinAPIError(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
   
   sBuff = String$(256, 0)
   lCount = FormatMessage( _
      FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
      0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
   If lCount Then
      WinAPIError = Left$(sBuff, lCount)
   End If
   
End Function


Public Property Get LoadResPicture(ByVal ID As Variant, ByVal Format As Variant) As IPicture
Dim sFile As String
Dim b() As Byte
Dim iFile As Integer
On Error GoTo ErrorHandler
   b = LoadResData(ID, Format)
   sFile = TempFileName("LRP")
   iFile = FreeFile
   Open sFile For Binary Access Write Lock Read As #iFile
   Put #iFile, , b
   Close #iFile
   iFile = 0
   Set LoadResPicture = LoadPicture(sFile)
   KillFile sFile
   Exit Property
ErrorHandler:
Dim lErr As Long, sErr As String
   lErr = Err.Number:   sErr = Err.Description
   If Not iFile = 0 Then Close #iFile
   KillFile sFile
   Err.Raise Err.Number, App.EXEName & ".cLoadResPicture", Err.Description
   Exit Property
End Property

Private Sub KillFile(ByVal sFile As String)
   On Error Resume Next
   Kill sFile
End Sub

Public Sub Resize()
UserControl_Resize
End Sub
Public Property Let ImageList(ByVal zImageList As ImageList)
    Set zImg = zImageList
End Property

Public Property Let Icon(ByVal zPicture As Integer)
    Set Logo.Picture = zImg.ListImages.Item(zPicture).Picture
End Property
