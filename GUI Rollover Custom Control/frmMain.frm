VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "GUI Button Example - Windows Media 11 Player GUI Exercise"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowMouseMovement 
      Caption         =   "Show Mouse Movement Feedback"
      Height          =   210
      Left            =   1530
      TabIndex        =   15
      Top             =   4230
      Value           =   1  'Checked
      Width           =   2745
   End
   Begin VB.CommandButton cmdReset_List 
      Caption         =   "&Reset List"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   1230
   End
   Begin VB.ListBox GUIButtonFeedback 
      Height          =   2205
      Left            =   90
      TabIndex        =   12
      Top             =   1770
      Width           =   5775
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   7995
      TabIndex        =   11
      Top             =   4080
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "GUI Button Configuration"
      Height          =   2385
      Left            =   6015
      TabIndex        =   3
      Top             =   1515
      Width           =   3180
      Begin EnterLeave.GUI_Rollover GUI_Rollover1 
         Height          =   435
         Left            =   195
         TabIndex        =   5
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   767
         Enabled         =   0   'False
         ImageNormal     =   "frmMain.frx":290FA
         ImageDisabled   =   "frmMain.frx":2A2F8
         ImageMask       =   "frmMain.frx":2B4F6
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Play Button Selectable"
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   1140
         MaskColor       =   &H80000012&
         TabIndex        =   10
         Top             =   1305
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable Play Button"
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   1140
         MaskColor       =   &H80000012&
         TabIndex        =   8
         Top             =   1005
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Enable Next Button"
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   1170
         MaskColor       =   &H80000012&
         TabIndex        =   6
         Top             =   1905
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Enable Previous Button"
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   1140
         MaskColor       =   &H80000012&
         TabIndex        =   4
         Top             =   435
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin EnterLeave.GUI_Rollover GUI_Rollover2 
         Height          =   435
         Left            =   225
         TabIndex        =   7
         Top             =   1860
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   767
         Enabled         =   0   'False
         ImageNormal     =   "frmMain.frx":2C6F4
         ImageDisabled   =   "frmMain.frx":2D8F2
         ImageMask       =   "frmMain.frx":2EAF0
      End
      Begin EnterLeave.GUI_Rollover GUI_Rollover3 
         Height          =   750
         Left            =   210
         TabIndex        =   9
         Top             =   930
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         Enabled         =   0   'False
         ImageNormal     =   "frmMain.frx":2FCEE
         ImageDisabled   =   "frmMain.frx":31AF0
         ImageMask       =   "frmMain.frx":338F2
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   885
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   870
         Width           =   950
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   510
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   950
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   510
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   950
      End
   End
   Begin EnterLeave.GUI_Rollover grPlay 
      Height          =   750
      Left            =   4275
      TabIndex        =   2
      Top             =   465
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      ImageNormal     =   "frmMain.frx":356F4
      ImageMask       =   "frmMain.frx":374F6
   End
   Begin EnterLeave.GUI_Rollover grPrevious 
      Height          =   390
      Left            =   3570
      TabIndex        =   0
      Top             =   630
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   688
   End
   Begin EnterLeave.GUI_Rollover grNext 
      Height          =   420
      Left            =   4980
      TabIndex        =   1
      Top             =   630
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   741
   End
   Begin VB.Label Label1 
      Caption         =   "Mouse Feedback"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   150
      TabIndex        =   13
      Top             =   1530
      Width           =   2115
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// Load Form and define buttons and settings
Private Sub Form_Load()
Dim ImagePath As String
    ImagePath = App.Path & "\Images\"
    
    
    'If chkShowMouseMovement Then MsgBox "Hello"
    
    '// Define Images for Play Button
    With grPlay
        .selectable = True
        .selected = False
        .ImageDisabled = LoadPicture(ImagePath & "PlayButton_Disabled.bmp")
        .ImageDown = LoadPicture(ImagePath & "PlayButton_Down.bmp")
        .ImageHover = LoadPicture(ImagePath & "PlayButton_Hover.bmp")
        .ImageMask = LoadPicture(ImagePath & "PlayButton_Mask.bmp")
        .ImageNormal = LoadPicture(ImagePath & "PlayButton_Normal.bmp")
        .ImageSelected = LoadPicture(ImagePath & "PlayButton_Selected.bmp")
        .ImageSelectedHover = LoadPicture(ImagePath & "PlayButton_SelectedHover.bmp")
    End With
    
    '// Define Images for Previous Button
        With grPrevious
        .selectable = False
        .selected = False
        .ImageDisabled = LoadPicture(ImagePath & "Previous_Disabled.bmp")
        .ImageDown = LoadPicture(ImagePath & "Previous_Down.bmp")
        .ImageHover = LoadPicture(ImagePath & "Previous_Hover.bmp")
        .ImageMask = LoadPicture(ImagePath & "Previous_Mask.bmp")
        .ImageNormal = LoadPicture(ImagePath & "Previous_Normal.bmp")
    End With
    
    '// Define Images for Next Button
    With grNext
        .selectable = False
        .selected = False
        .ImageDisabled = LoadPicture(ImagePath & "Next_Disabled.bmp")
        .ImageDown = LoadPicture(ImagePath & "Next_Down.bmp")
        .ImageHover = LoadPicture(ImagePath & "Next_Hover.bmp")
        .ImageMask = LoadPicture(ImagePath & "Next_Mask.bmp")
        .ImageNormal = LoadPicture(ImagePath & "Next_Normal.bmp")
    End With
End Sub


'// Configuration Checkboxes
Private Sub Check1_Click()
    grPlay.Enabled = IIf(Check1.value = 1, True, False)
End Sub

Private Sub Check2_Click()
    grPrevious.Enabled = IIf(Check2.value = 1, True, False)
End Sub

Private Sub Check3_Click()
    grNext.Enabled = IIf(Check3.value = 1, True, False)
End Sub

Private Sub Check5_Click()
    grPlay.selectable = IIf(Check5.value = 1, True, False)
End Sub


'// Command Buttons
Private Sub cmdReset_List_Click()
    GUIButtonFeedback.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'//grPrevious Button Events
Private Sub grPrevious_OnMouseClick()
    GUIButtonFeedback.AddItem "Previous Button: Clicked", 0
End Sub

Private Sub grPrevious_OnMouseEnter()
    GUIButtonFeedback.AddItem "Previous Button: Mouse Enter", 0
End Sub

Private Sub grPrevious_OnMouseLeave()
    GUIButtonFeedback.AddItem "Previous Button: Mouse Leave", 0
End Sub

Private Sub grPrevious_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkShowMouseMovement Then
        GUIButtonFeedback.AddItem "Previous Button: Mouse Move (X= " & X & "; Y= " & Y & ")", 0
    End If
End Sub

Private Sub grPrevious_onMouseSelectionChange()
    GUIButtonFeedback.AddItem "Previous Button: Selection State Changed to " & IIf(grPrevious.selected = True, "Selected", "Not Selected"), 0
End Sub


'//grPlay Button Events
Private Sub grPlay_OnMouseClick()
    GUIButtonFeedback.AddItem "Previous Button: Clicked", 0
End Sub

Private Sub grPlay_OnMouseEnter()
    GUIButtonFeedback.AddItem "Previous Button: Mouse Enter", 0
End Sub

Private Sub grPlay_OnMouseLeave()
    GUIButtonFeedback.AddItem "Previous Button: Mouse Leave", 0
End Sub

Private Sub grPlay_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkShowMouseMovement Then
        GUIButtonFeedback.AddItem "Previous Button: Mouse Move (X= " & X & "; Y= " & Y & ")", 0
    End If
End Sub

Private Sub grPlay_onMouseSelectionChange()
    GUIButtonFeedback.AddItem "Play Button: Selection State Changed to " & IIf(grPlay.selected = True, "Selected", "Not Selected"), 0
End Sub


'//grNext Button Events
Private Sub grNext_OnMouseClick()
    GUIButtonFeedback.AddItem "Next Button: Clicked", 0
End Sub

Private Sub grNext_OnMouseEnter()
    GUIButtonFeedback.AddItem "Next Button: Mouse Enter", 0
End Sub

Private Sub grNext_OnMouseLeave()
    GUIButtonFeedback.AddItem "Next Button: Mouse Leave", 0
End Sub

Private Sub grNext_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkShowMouseMovement Then
        GUIButtonFeedback.AddItem "Next Button: Mouse Move (X= " & X & "; Y= " & Y & ")", 0
    End If
End Sub

Private Sub grNext_onMouseSelectionChange()
    GUIButtonFeedback.AddItem "Next Button: Selection State Changed to " & IIf(grNex.selected = True, "Selected", "Not Selected"), 0
End Sub

