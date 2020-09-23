VERSION 5.00
Begin VB.Form frmTestMsgBoxEx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MsgboxEx"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5205
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Height          =   5160
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   5895
      Begin VB.CheckBox chk 
         Caption         =   "Cancel 6"
         Height          =   255
         Index           =   26
         Left            =   4560
         TabIndex        =   73
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Cancel 5"
         Height          =   255
         Index           =   25
         Left            =   4560
         TabIndex        =   72
         Top             =   3180
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Cancel 4"
         Height          =   255
         Index           =   24
         Left            =   4560
         TabIndex        =   71
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Cancel 3"
         Height          =   255
         Index           =   23
         Left            =   4560
         TabIndex        =   70
         Top             =   2770
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Cancel 2"
         Height          =   255
         Index           =   22
         Left            =   4560
         TabIndex        =   69
         Top             =   2540
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Cancel 1"
         Height          =   255
         Index           =   21
         Left            =   4560
         TabIndex        =   68
         Top             =   2310
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Default 6"
         Height          =   255
         Index           =   18
         Left            =   4560
         TabIndex        =   67
         Top             =   2080
         Width           =   1095
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   42
         Top             =   4320
         Width           =   2535
         Begin VB.CommandButton cmd 
            Caption         =   "Show"
            Height          =   375
            Index           =   15
            Left            =   0
            TabIndex        =   45
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optDialog 
            Caption         =   "InputBox"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   44
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optDialog 
            Caption         =   "Msgbox"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.OptionButton opt 
         Caption         =   "Custom"
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   41
         Top             =   4800
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "RightAbove"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   40
         Top             =   4560
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "RightCenter"
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   39
         Top             =   4320
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "RightBelow"
         Height          =   255
         Index           =   6
         Left            =   4560
         TabIndex        =   38
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "LeftAbove"
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   37
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "LeftCenter"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   36
         Top             =   4800
         Width           =   1335
      End
      Begin VB.OptionButton opt 
         Caption         =   "LeftBelow"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   35
         Top             =   4560
         Width           =   1335
      End
      Begin VB.OptionButton opt 
         Caption         =   "CenterBelow"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   34
         Top             =   4320
         Width           =   1335
      End
      Begin VB.OptionButton opt 
         Caption         =   "CenterAbove"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   33
         Top             =   4080
         Width           =   1335
      End
      Begin VB.OptionButton opt 
         Caption         =   "CenterCenter"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   32
         Top             =   3840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use My Icon"
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   4695
         MousePointer    =   3  'I-Beam
         TabIndex        =   29
         Text            =   "10"
         Top             =   3555
         Width           =   375
      End
      Begin VB.CheckBox chk 
         Caption         =   "Default 5"
         Height          =   255
         Index           =   20
         Left            =   4560
         TabIndex        =   28
         Top             =   1850
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Default 4"
         Height          =   255
         Index           =   19
         Left            =   4560
         TabIndex        =   27
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Default 3"
         Height          =   255
         Index           =   17
         Left            =   4560
         TabIndex        =   26
         Top             =   1390
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Default 2"
         Height          =   255
         Index           =   16
         Left            =   4560
         TabIndex        =   25
         Top             =   1160
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Default 1"
         Height          =   255
         Index           =   15
         Left            =   4560
         TabIndex        =   24
         Top             =   930
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "No Move"
         Height          =   255
         Index           =   14
         Left            =   4560
         TabIndex        =   23
         Top             =   700
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Cancel Err"
         Height          =   255
         Index           =   13
         Left            =   4560
         TabIndex        =   22
         Top             =   470
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Hide Timer"
         Height          =   255
         Index           =   12
         Left            =   4560
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "TopMost"
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   20
         Top             =   3210
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Beep"
         Height          =   255
         Index           =   10
         Left            =   2760
         TabIndex        =   19
         Top             =   2940
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Information Icon"
         Height          =   255
         Index           =   9
         Left            =   2745
         TabIndex        =   18
         Top             =   2670
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Exclamation Icon"
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   17
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Question Icon"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   16
         Top             =   2130
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Critical Icon"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   15
         Top             =   1860
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Retry/Cancel"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   14
         Top             =   1590
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Yes/No"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Yes/No/Cancel"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   12
         Top             =   1050
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Abort/Retiry/Ignore"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   11
         Top             =   780
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "OK/Cancel"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Top             =   510
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "OK Only"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   2895
         Index           =   1
         Left            =   120
         MousePointer    =   3  'I-Beam
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "TestMsgBoxEx.frx":0000
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   7
         Text            =   "Type the title"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Type a message or copy some RTF code."
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Enter a title:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Timer:           seconds"
         Height          =   255
         Left            =   4200
         TabIndex        =   30
         Top             =   3600
         Width           =   1575
      End
   End
   Begin VB.Frame fra 
      Height          =   2955
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmd 
         Caption         =   "Test 8"
         Height          =   495
         Index           =   11
         Left            =   2040
         TabIndex        =   53
         Top             =   1260
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 7"
         Height          =   495
         Index           =   10
         Left            =   2040
         TabIndex        =   52
         Top             =   750
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 6"
         Height          =   495
         Index           =   9
         Left            =   2040
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 9"
         Height          =   495
         Index           =   12
         Left            =   2040
         TabIndex        =   50
         Top             =   1770
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 10"
         Height          =   495
         Index           =   13
         Left            =   2040
         TabIndex        =   49
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 5"
         Height          =   495
         Index           =   8
         Left            =   1080
         TabIndex        =   58
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 4"
         Height          =   495
         Index           =   7
         Left            =   1080
         TabIndex        =   57
         Top             =   1770
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 3"
         Height          =   495
         Index           =   6
         Left            =   1080
         TabIndex        =   56
         Top             =   1260
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 2"
         Height          =   495
         Index           =   5
         Left            =   1080
         TabIndex        =   55
         Top             =   750
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Test 1"
         Height          =   495
         Index           =   4
         Left            =   1080
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "About"
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   63
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Custom Buttons"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   62
         Top             =   1770
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Modeless"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   1260
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Modal"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   750
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "InputBoxEx"
         Height          =   495
         Index           =   14
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Global Attributes"
      Height          =   2160
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   3
         ItemData        =   "TestMsgBoxEx.frx":0038
         Left            =   60
         List            =   "TestMsgBoxEx.frx":0054
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1140
         Width           =   1995
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   2
         ItemData        =   "TestMsgBoxEx.frx":009E
         Left            =   60
         List            =   "TestMsgBoxEx.frx":00BD
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   840
         Width           =   1995
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   1
         ItemData        =   "TestMsgBoxEx.frx":0137
         Left            =   60
         List            =   "TestMsgBoxEx.frx":0144
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Gradient"
         Height          =   375
         Index           =   1
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkGlobal 
         Caption         =   "Show Focus Rect"
         Height          =   255
         Index           =   2
         Left            =   1500
         TabIndex        =   5
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkGlobal 
         Caption         =   "Show Divider"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   0
         ItemData        =   "TestMsgBoxEx.frx":0167
         Left            =   60
         List            =   "TestMsgBoxEx.frx":017A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "BackColor"
         Height          =   375
         Index           =   0
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTestMsgBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements iRichDialogParent

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Private Const HWND_TOPMOST = -1
    'Private Const HWND_NOTOPMOST = -2
    Private Const SWP_NOSIZE = &H1
    Private Const SWP_SHOWWINDOW = &H40
    Private Const SWP_NOACTIVATE = &H10
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOZORDER = &H4
    Private Const SWP_NOOWNERZORDER = &H200
    Private Const SWP_NOREDRAW = &H8

Private mbShowResult As Boolean
Const ShowingResultTag = 98457

Private Sub ShowResult(piResult As eRichDialogReturn)
    Dim lsTemp As String
    Select Case piResult
        Case rdOK
            lsTemp = "OK"
        Case rdCancel
            lsTemp = "Cancel"
        Case rdAbort
            lsTemp = "Abort"
        Case rdRetry
            lsTemp = "Retry"
        Case rdIgnore
            lsTemp = "Ignore"
        Case rdYes
            lsTemp = "Yes"
        Case rdNo
            lsTemp = "No"
        Case rdButton1
            lsTemp = "Button 1"
        Case rdButton2
            lsTemp = "Button 2"
        Case rdButton3
            lsTemp = "Button 3"
        Case rdButton4
            lsTemp = "Button 4"
        Case rdButton5
            lsTemp = "Button 5"
        Case rdButton6
            lsTemp = "Button 6"
        Case rdAutoAnswer
            lsTemp = "Auto Answer"
        'case rdFormUnloading
            'Exit Sub
        Case Else
            'Debug.Assert False
    End Select
    MsgBoxEx lsTemp, rdInformation, , , , , , , ShowingResultTag, Me
End Sub

Private Sub chkGlobal_Click(Index As Integer)
    With RichDialogGUIDefaults
        Select Case Index
            Case 1
                .ShowDivider = chkGlobal(Index).Value = 1
            Case 2
                .ShowFocusRect = chkGlobal(Index).Value = 1
        End Select
    End With
End Sub


Private Sub cmb_Click(Index As Integer)
    With RichDialogGUIDefaults
        Select Case Index
            Case 0
                .GradientType = cmb(Index).ListIndex
            Case 1
                .CaptionEffect = cmb(Index).ListIndex
            Case 2
                .FontStyle = cmb(Index).ItemData(cmb(Index).ListIndex)
            Case 3
                .ButtonDrawStyle = cmb(Index).ListIndex
        End Select
    End With
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim liAnswer As eRichDialogReturn
    Dim lsMsg As String
    Dim liAtt As eRichDialogAttributes
    Dim lsTitle As String
    Dim liTime As Long
    Dim lvButtons
    Dim liIcon As Long
    Dim liHwnd As Long
    Select Case Index
        Case 0
            lsMsg = "This messagebox is modal."
            liAtt = rdOK + rdInformation
            lsTitle = "Modal"
        Case 1
            MsgBoxEx "This messagebox is modeless.", rdInformation, "Modeless", , , , , , , Me, True
            Exit Sub
        Case 2
            lsMsg = "This messagebox has custom buttons, by providing a variant array as an optional argument." & vbCrLf & "The Default and Cancel button are the fourth button."
            liAtt = rdCancelButton4 + rdDefaultButton4
            lsTitle = "Custom Buttons"
            lvButtons = Array("One", "Two", "Three", "Four")
        Case 3
            lsMsg = LoadResString(101)
            lsTitle = "About MsgBoxEx"
        Case 4
            lsMsg = "This messagebox is the default.  Only a caption is provided."
        Case 5
            lsMsg = "This messagebox has OK/Cancel buttons. Both the Default and Cancel Button are the first button."
            liAtt = rdOKCancel + rdDefaultButton1 + rdCancelButton1
        Case 6
            lsMsg = "This messagebox has Abort/Retry/Ignore buttons.  Both the default and Cancel button are the second button.  The icon is the standard Question Icon"
            liAtt = rdAbortRetryIgnore + rdDefaultButton2 + rdCancelButton2 + rdQuestion
        Case 7
            lsMsg = "This messagebox has Yes/No/Cancel Buttons. Both the Default and Cancel button are the third button. The icon is the standard information icon."
            liAtt = rdYesNoCancel + rdDefaultButton3 + rdCancelButton3 + rdInformation
        Case 8
            lsMsg = "This messagebox has Yes/No Buttons and the Exclamation standard icon."
            liAtt = rdYesNo + rdExclamation
        Case 9
            lsMsg = "This messagebox has Retry/Cancel buttons and the Critical standard icon."
            liAtt = rdRetryCancel + rdCritical
        Case 10
            lsMsg = "This messagebox has the icon of frmTestMsgBoxEx as it's icon"
            lsTitle = "Custom Icon"
            liIcon = Icon.Handle
        Case 11
            lsMsg = "This messagebox times out after 7 seconds."
            liTime = 7
        Case 12
            lsMsg = "This messagebox times out after 7 seconds, but counts down silently."
            liAtt = rdHideTimeOutCountdown
            liTime = 7
        Case 13
            lsMsg = "This messagebox is centered over the form."
            liHwnd = Me.hWnd
        Case 14
            Dim lsAnswer As String
            On Error Resume Next
            lsAnswer = InputboxEx("This is a modal Input Box", rdCancelRaiseError + rdOKCancel, "Input", "default goes here", Me.Icon.Handle, True)
            If Err.Number = 20001 Then lsAnswer = "(Canceled)"
            MsgBoxEx lsAnswer
            Exit Sub
        Case 15
            ShowCustom
            Exit Sub
    End Select
    liAnswer = MsgBoxEx(lsMsg, liAtt, lsTitle, liIcon, liHwnd, , lvButtons, liTime, , Me)
End Sub

Private Sub ShowCustom()
    Dim liAttributes As eRichDialogAttributes
    Dim liPosition As Long
    Dim i As Long
    Dim lvTemp
    
    For Each lvTemp In chk
        Select Case lvTemp.Index
            Case Is <= rdRetryCancel
                If lvTemp.Value = 1 Then liAttributes = liAttributes + lvTemp.Index
            Case Else
                If lvTemp.Value = 1 Then liAttributes = liAttributes + 2 ^ (lvTemp.Index - 2)
        End Select
        
    Next

    For i = 0 To opt.UBound
        If opt(i).Value Then
            liPosition = i
            Exit For
        End If
    Next
    If optDialog(1).Value Then
        InputboxEx txt(1).Text, liAttributes, txt(0).Text, "Default Val", IIf(Check1.Value = 1, Icon.Handle, 0), hWnd, liPosition, , Val(txt(2).Text), , Me
    Else
        MsgBoxEx txt(1).Text, liAttributes, txt(0).Text, IIf(Check1.Value = 1, Icon.Handle, 0), hWnd, liPosition, , Val(txt(2).Text), , Me
    End If
    
End Sub

Private Sub cmdColor_Click(Index As Integer)
    Dim Col As SelectedColor
    Col = ShowColor(hWnd)
    If Not Col.bCanceled Then
        cmdColor(Index).BackColor = Col.oSelectedColor
        If Index = 0 Then RichDialogGUIDefaults.BackColor = Col.oSelectedColor Else RichDialogGUIDefaults.GradientColor = Col.oSelectedColor
    End If
End Sub

Private Sub Form_Initialize()
    Dim i As Long
    mbShowResult = True
    With cmb
        For i = .LBound To .UBound
            .Item(i).ListIndex = 0
        Next
    End With
End Sub

Private Sub iRichDialogParent_HasReturned(ByVal Dialog As RichDialogs.iRichDialog)
    On Error Resume Next
    With Dialog.Info
        If .Tag <> ShowingResultTag Then
            If mbShowResult Then ShowResult .Answer
        Else
            mbShowResult = Not .CheckBoxValue
        End If
    End With
End Sub

Private Sub iRichDialogParent_QueryInfo(ByVal Dialog As RichDialogs.iRichDialog, bCancel As Boolean)

    With Dialog.Info
        If .Position = rdCustom Then
            .Message = .Message & vbCrLf & "This message will be placed even with the right and left edges of this form, and will not be moveable."
            .Attributes = .Attributes Or rdDisallowMove
        End If
        If .Tag = ShowingResultTag Then
            .CheckBoxStatement = "Stop showing my answers"
        End If
    End With

End Sub

Private Sub iRichDialogParent_WillShow(ByVal Dialog As RichDialogs.iRichDialog, bCancel As Boolean)
    If Dialog.Info.Position = rdCustom Then SetWindowPos Dialog.hWnd, 0, ScaleX(left, vbTwips, vbPixels), ScaleY(top, vbTwips, vbPixels), 0, 0, SWP_NOACTIVATE + SWP_NOSIZE + SWP_NOZORDER
End Sub

