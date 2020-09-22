VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Application - document.docx"
   ClientHeight    =   5280
   ClientLeft      =   -45
   ClientTop       =   -405
   ClientWidth     =   8520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ACPRibbon ACPRibbon1 
      Height          =   2130
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   3757
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":052F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0686
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1522
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":210C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26BF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "All icons are here!"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Sorry for my poor English, i´m Brazilian."
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":2C5A
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   8295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":2CE1
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub Form_Load()

'# SET Theme
ACPRibbon1.Theme = 2    ' 0 - Black
                        ' 1 - Blue
                        ' 2 - Silver
                        

'# Set ImageList to use for icons
ACPRibbon1.ImageList = ImageList1

'# Set Circle Menu Button Picture with Index of image on imagelist
ACPRibbon1.Icon = 1

'# Show Caption of Form
ACPRibbon1.Caption = Me.Caption

'# Show Button to Customize Menu
ACPRibbon1.ShowCustomMenu = True

'# Add TopButtons ---   ID - Capt. - Icons
ACPRibbon1.AddTopButton "1", "New", 2
ACPRibbon1.AddTopButton "2", "Open", 3
ACPRibbon1.AddTopButton "3", "Print", 4
ACPRibbon1.AddTopButton "3", "Save", 5

'# Add Tabs ---   ID - Caption
ACPRibbon1.AddTab "1", "Tab 1"
ACPRibbon1.AddTab "2", "Tab 2"
ACPRibbon1.AddTab "3", "Sample Tab"
ACPRibbon1.AddTab "4", "New Tab"
ACPRibbon1.AddTab "5", "WOW"

'# Add Cats ---   ID - Tab - Caption - ShowDialogButton
ACPRibbon1.AddCat "1", "1", "Group 1", False
ACPRibbon1.AddCat "2", "1", "One very large group", True
ACPRibbon1.AddCat "3", "1", "Test", True
ACPRibbon1.AddCat "4", "2", "More one group", True
ACPRibbon1.AddCat "5", "2", "Hi!", False
ACPRibbon1.AddCat "6", "3", "Hello World!", False

'# Add Button ---    ID - Cat - Capt. - Icons -   More Arrow   - ToolTip
ACPRibbon1.AddButton "1", "1", "Table", 6, False, "Insert a new Table"
ACPRibbon1.AddButton "2", "1", "Insert Picture", 7
ACPRibbon1.AddButton "2", "1", "Insert" & vbNewLine & "Picture", 7
ACPRibbon1.AddButton "3", "2", "Graph", 8
ACPRibbon1.AddButton "4", "2", "Graph", 8, True
ACPRibbon1.AddButton "5", "3", "Clip Art", 9
ACPRibbon1.AddButton "6", "4", "SmartDraw", 10

'# Repaint Ribbon
ACPRibbon1.Refresh

End Sub

Private Sub Form_Resize()

'# this procedure will resize the ribbon
ACPRibbon1.Resize

End Sub

Private Sub ACPRibbon1_MainMenuClick()

'# This Event occurs on click in Main Button Menu
MsgBox "Main Menu Click"

End Sub

Private Sub ACPRibbon1_CustomClick()

'# This Event occurs on click in Custom Button Menu
MsgBox "Custom Click"

End Sub

Private Sub ACPRibbon1_MenuClick(ByVal ID As String, ByVal Caption As String)

'# This Event occurs when click on each Menu Button
MsgBox "MenuClick: " & ID & "--" & Caption

End Sub

Private Sub ACPRibbon1_TabClick(ByVal ID As String, ByVal Caption As String)

'# This Event occurs when click on each tab
MsgBox "TabSelected: " & ID & "--" & Caption

End Sub

Private Sub ACPRibbon1_CatClick(ByVal ID As String, ByVal Caption As String)

'# This Event occurs when click on each ShowDialogButton for each Categorie
MsgBox "ShowDialogClick: " & ID & "--" & Caption

End Sub

Private Sub ACPRibbon1_ButtonClick(ByVal ID As String, ByVal Caption As String)

'# This Event occurs when click on each Button
MsgBox "ButtonClick: " & ID & "--" & Caption

End Sub


