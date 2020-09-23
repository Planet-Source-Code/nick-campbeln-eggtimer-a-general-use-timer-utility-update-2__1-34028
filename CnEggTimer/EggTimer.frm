VERSION 5.00
Begin VB.Form frmEggTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EggTimer"
   ClientHeight    =   2355
   ClientLeft      =   11220
   ClientTop       =   10170
   ClientWidth     =   3960
   Icon            =   "EggTimer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAbout 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3600
      TabIndex        =   25
      Top             =   0
      Width           =   495
      Begin VB.Image imgAbout 
         Height          =   240
         Left            =   70
         Picture         =   "EggTimer.frx":0442
         ToolTipText     =   "About Cn's EggTimer"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.OptionButton optType 
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
      Index           =   2
      Left            =   1880
      TabIndex        =   21
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optType 
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
      Index           =   1
      Left            =   2050
      TabIndex        =   20
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton optType 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   1660
      TabIndex        =   19
      Top             =   1080
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.Frame fraDingAt 
      Caption         =   " Ding At "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   3735
      Begin VB.ComboBox cboTimeHour 
         Height          =   315
         ItemData        =   "EggTimer.frx":058C
         Left            =   240
         List            =   "EggTimer.frx":05B7
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboTimeMinute 
         Height          =   315
         ItemData        =   "EggTimer.frx":05E2
         Left            =   1440
         List            =   "EggTimer.frx":069A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboTimeAMPM 
         Height          =   315
         ItemData        =   "EggTimer.frx":078E
         Left            =   2640
         List            =   "EggTimer.frx":0798
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblTimeHours 
         Caption         =   "Hour"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblTimeMinutes 
         Caption         =   "Minute"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblTimeAMPM 
         Caption         =   "AM / PM"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraTimeRemaining 
      Caption         =   " Time Remaining "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   3735
      Begin VB.Label lblTimeRemaining 
         Alignment       =   2  'Center
         Caption         =   "0:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1620
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   4920
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2580
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame fraDuration 
      Caption         =   " Duration "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3735
      Begin VB.ComboBox cboSeconds 
         Height          =   315
         ItemData        =   "EggTimer.frx":07A4
         Left            =   2640
         List            =   "EggTimer.frx":085C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboMinutes 
         Height          =   315
         ItemData        =   "EggTimer.frx":0950
         Left            =   1440
         List            =   "EggTimer.frx":0A08
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboHours 
         Height          =   315
         ItemData        =   "EggTimer.frx":0AFC
         Left            =   240
         List            =   "EggTimer.frx":0B59
         TabIndex        =   5
         Text            =   "1"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblSeconds 
         Caption         =   "Seconds"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblMinutes 
         Caption         =   "Minutes"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblHours 
         Caption         =   "Hours"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   660
      TabIndex        =   0
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblOption 
      AutoSize        =   -1  'True
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   870
      TabIndex        =   24
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblOption 
      AutoSize        =   -1  'True
      Caption         =   "Ding At"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2390
      TabIndex        =   23
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label lblOption 
      AutoSize        =   -1  'True
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1747
      TabIndex        =   22
      Top             =   1590
      Width           =   480
   End
   Begin VB.Image imgPause 
      Height          =   240
      Left            =   480
      Picture         =   "EggTimer.frx":0BB7
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgTimer 
      Height          =   240
      Left            =   120
      Picture         =   "EggTimer.frx":0D01
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmEggTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#####################################################################################
'#  EggTimer - A General Use Timer Utility
'#      By: Nick Campbeln
'#
'#  Revision Histroy:
'#      1.1 (Apr 23, 2002):
'#          Added 'Timer' option
'#          Fully commented code
'#          Added shameless self promotion
'#      1.0 (Apr 22, 2002):
'#          Initial Release
'#
'#      Copyright © 2002 Nick Campbeln (opensource@nick.campbeln.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  Please remember to vote on PSC.com if you like this code!
'#  Code URL: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=34028&lngWId=1
'#####################################################################################

    '#### Mouse click constants for the TrayIcon Class
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206

    '#### Setup the necessary global vars
Private cTrayIcon As New clsTrayIcon
Private sTimeRemaining As String
Private lHours As Long
Private lMinutes As Long
Private lSeconds As Long



'#####################################################################################
'# Form Subs/Functions
'#####################################################################################
'#########################################################
'# Form_Load() inits the form elements with their initial values
'#########################################################
Private Sub Form_Load()
    Dim iHour As Integer

        '#### Set the default value for cboMinutes, cboSeconds, cboTimeMinute and determine the hour
    cboMinutes.ListIndex = 0
    cboSeconds.ListIndex = 0
    cboTimeMinute.ListIndex = 0
    iHour = Hour(DateAdd("h", 1, Now))

        '#### Set the default value for cboTimeAMPM
    If (iHour >= 12) Then
        cboTimeAMPM.ListIndex = 1
        iHour = iHour - 12
    Else
        cboTimeAMPM.ListIndex = 0
    End If

        '#### Set the default value for cboTimeHour
    Select Case iHour
        Case 12, 0
            cboTimeHour.ListIndex = cboTimeHour.ListCount - 1
        Case Else
            cboTimeHour.ListIndex = iHour - 1
    End Select
End Sub


'#########################################################
'# Form_QueryUnload() ensures that the user really wants to quick if the timer is currently running
'#########################################################
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lReturn As Long

        '#### If the timer is currently enabled
    If (tmrTimer.Enabled) Then
            '#### Reset the cTrayIcon with the paused information as the app does not run while a msgbox() is popped and pop the prompt
        Call cTrayIcon.Modify(imgPause, sTimeRemaining & " (PAUSED)")
        lReturn = MsgBox("This will stop your timer, do you still wish to exit?" & vbCrLf & vbCrLf & "NOTE: Your timer is currently paused.", vbYesNo + vbQuestion, "Stop Timer?")

            '#### If the user decided to exit, do so now
        If (lReturn = vbYes) Then
            Set frmEggTimer = Nothing

            '#### Else cancel the unload and reset the tray icon
        Else
            Call cTrayIcon.Modify(imgTimer, sTimeRemaining)
            Cancel = 1
        End If
    End If
End Sub


'#########################################################
'# Form_Unload() correctly and completly destroys the objects
'#########################################################
Private Sub Form_Unload(Cancel As Integer)
        '#### Explicitly destroy cTrayIcon and the main form itself
    Set cTrayIcon = Nothing
    Set frmEggTimer = Nothing
End Sub


'#########################################################
'# Form_Resize() ensures that the correctly disapears when the user minimizes it
'#########################################################
Private Sub Form_Resize()
        '#### If the timer is running and the form was just minimized, hide it completly
    If (tmrTimer And frmEggTimer.WindowState = 1) Then
        frmEggTimer.Visible = False
    End If
End Sub


'#########################################################
'# Form_MouseMove() acts as the event processor for the cTrayIcon
'#########################################################
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '#### If the user clicked using ANY button in the cTrayIcon
    Select Case (X / Screen.TwipsPerPixelX)
        Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MBUTTONDBLCLK, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_RBUTTONDBLCLK
                '#### Show and SetFocus to the form, resetting its WindowState property so that the Form_Resize() sub does the right thing
            frmEggTimer.WindowState = 0
            frmEggTimer.Visible = True
            frmEggTimer.SetFocus
    End Select
End Sub



'#####################################################################################
'# Control Subs/Functions
'#####################################################################################
'#########################################################
'# imgAbout_Click() enables my shameless self promotion (please leave intact =)
'#########################################################
Private Sub imgAbout_Click()
    Call MsgBox("Cn's EggTimer v" & App.Major & "." & App.Minor & "." & App.Revision & " was created by Nick Campbeln." & vbCrLf & vbCrLf & "For the source to this and other applications, please" & vbCrLf & "visit http://software.isontheweb.com/", vbOKOnly + vbInformation + vbApplicationModal, "About Cn's EggTimer")
End Sub


'#########################################################
'# lblOption_Click() and optType_Click() enables the user to select the type of timer
'#########################################################
Private Sub lblOption_Click(Index As Integer)
        '#### Foward the call onto the option's sub
    Call optType_Click(Index)
    optType(Index).Value = True
End Sub
Private Sub optType_Click(Index As Integer)
        '#### Determine the option the user selectd and process accordingly
    Select Case Index
            '#### If the user selected the 'duration' option
        Case 0
                '#### Move fraDuration into view and enable it move fraDingAt out of the way and disable it, and reset fraTimeRemaining's caption
            fraDuration.Top = 0
            fraDuration.Enabled = True
            fraDingAt.Top = frmEggTimer.Height * 2
            fraDingAt.Enabled = False
            fraTimeRemaining.Caption = " Time Remaining "

            '#### Else if the user selected the 'ding at' option
        Case 1
                '#### Move fraDuration out of the way and disable it, move fraDingAt into view and enable it and reset fraTimeRemaining's caption
            fraDuration.Top = frmEggTimer.Height * 2
            fraDuration.Enabled = False
            fraDingAt.Top = 0
            fraDingAt.Enabled = True
            fraTimeRemaining.Caption = " Time Remaining "

            '#### Else if the user selected the 'timer' option
        Case 2
                '#### Disable both of the other options frames and reset the caption of fraTimeRemaining
            fraDuration.Enabled = False
            fraDingAt.Enabled = False
            fraTimeRemaining.Caption = " Running Time "
    End Select
End Sub


'#########################################################
'# cmdStart_Click() inits lSeconds/lMinutes/lHours and starts the tmrTimer control
'#########################################################
Private Sub cmdStart_Click()
        '#### Make sure the value in cboHours is a number
    If (IsNumeric(cboHours.Text)) Then
            '#### If the user selected the 'duration' option
        If (optType(0).Value) Then
                '#### Set the global vars based on the user's inputs
            lHours = CLng(cboHours.Text)
            lMinutes = CLng(cboMinutes.Text)
            lSeconds = CLng(cboSeconds.Text)

            '#### Else if the user selected the 'ding at' option
        ElseIf (optType(1).Value) Then
                '#### Determine the lHours and lMinutes via DateDiff()
            lHours = DateDiff("h", Now, CDate(cboTimeHour.Text & ":" & cboTimeMinute.Text & ":00 " & cboTimeAMPM & " " & Date))
            lMinutes = DateDiff("n", Now, CDate(cboTimeHour.Text & ":" & cboTimeMinute.Text & ":00 " & cboTimeAMPM & " " & Date)) Mod 60

                '#### If the lHours is the current hour
            If (lHours = 0) Then
                    '#### If the lMinutes is the current minute
                If (lMinutes = 0) Then
                        '#### Set the lHours/lMinutes to the max time
                    lHours = 23
                    lMinutes = 60

                    '#### If the lMinutes are in the past
                ElseIf (lMinutes < 0) Then
                        '#### Set lHours to the max and calculate the remaining lMinutes
                    lHours = 23
                    lMinutes = 60 + lMinutes

                    '#### Else the lMinutes are in the future
                Else
                        '#### The value currently held in lHours/lMinutes are correct, so do nothing
                   'lHours = 0
                   'lMinutes = lMinutes
                End If

                '#### Else if the lHours is in the past
            ElseIf (lHours < 0) Then
                    '#### If the lMinutes is the current minute
                If (lMinutes = 0) Then
                        '#### Dec lHours from the max based on the number of hours in the past and set lMinutes to the max time
                    lHours = 23 + lHours
                    lMinutes = 60

                    '#### If the lMinutes are in the past
                ElseIf (lMinutes < 0) Then
                        '#### Dec lHours from the max based on the number of hours in the past and calculate the remaining lMinutes
                    lHours = 23 + lHours
                    lMinutes = 60 + lMinutes

                    '#### Else the lMinutes are in the future
                Else
                        '#### Dec lHours from the max based on the number of hours in the past and do nothing to lMinutes as it currently hold the correct value
                    lHours = 23 + lHours
                   'lMinutes = lMinutes
                End If

                '#### Else it is in the future
            Else
                    '#### If the lMinutes is the current minute
                If (lMinutes = 0) Then
                        '#### Dec lHours as it holds an extra hour and set lMinutes to the max time
                    lHours = lHours - 1
                    lMinutes = 60

                    '#### If the lMinutes are in the past
                ElseIf (lMinutes < 0) Then
                        '#### Dec lHours as it holds an extra hour and calculate the remaining lMinutes
                    lHours = lHours - 1
                    lMinutes = 60 + lMinutes

                    '#### Else the lMinutes are in the future
                Else
                        '#### Dec lHours as it holds an extra hour and do nothing to lMinutes as it currently hold the correct value
                    lHours = lHours - 1
                   'lMinutes = lMinutes
                End If
            End If

                '#### Determine the seconds, fixing the lMinutes if necessary
            lSeconds = Abs(Second(Now) - 60)
            If (lSeconds <> 0) Then lMinutes = lMinutes - 1

            '#### Else we'll assume the user chose the 'timer' option
        Else 'If (optType(2).Value) Then
                '#### Reset the globals to 0 as the timer option counts up
            lSeconds = 0
            lMinutes = 0
            lHours = 0
        End If

            '#### Enable/disable and set the captions of the buttons and the timer
        optType(0).Enabled = False
        optType(1).Enabled = False
        optType(2).Enabled = False
        cmdStart.Enabled = False
        cmdPause.Enabled = True
        cmdPause.Caption = "Pause"
        cmdStop.Enabled = True
        tmrTimer.Enabled = True

            '#### Add the cTrayIcon, move fraAbout and fraTimeRemaining over the top of the current frame and hide the form
        Call cTrayIcon.Add(frmEggTimer, imgTimer, sTimeRemaining)
        fraTimeRemaining.Top = 0
        fraTimeRemaining.ZOrder (0)
        fraAbout.ZOrder (0)
        frmEggTimer.Visible = False

        '#### Else the hours were not numeric, so prompt the user
    Else
        Call MsgBox("Please enter a numeric value for 'Hours'", vbOKOnly + vbCritical, "Invalid 'Hours' Setting")
    End If
End Sub


'#########################################################
'# cmdPause_Click() pauses/resumes the tmrTimer control based on the enabled state of the timer
'#########################################################
Private Sub cmdPause_Click()
        '#### Reset the timer and update the cTrayIcon
    tmrTimer.Enabled = CBool(Not tmrTimer.Enabled)

        '#### Determine the new value for the timer's Enabled property
    If (tmrTimer.Enabled) Then
            '#### Since the timer is enabled, set the caption to 'Pause' and set the cTrayIcon
        cmdPause.Caption = "Pause"
        Call cTrayIcon.Modify(imgTimer, sTimeRemaining)
    Else
            '#### Since the timer is disabled, set the caption to 'Resume' and set the cTrayIcon
        cmdPause.Caption = "Resume"
        Call cTrayIcon.Modify(imgPause, sTimeRemaining & " (PAUSED)")
    End If
End Sub


'#########################################################
'# cmdStop_Click() stops the tmrTimer control
'#########################################################
Private Sub cmdStop_Click()
        '#### Enable/disable the buttons and the timer
    optType(0).Enabled = True
    optType(1).Enabled = True
    optType(2).Enabled = True
    cmdStart.Enabled = True
    cmdPause.Enabled = False
    cmdStop.Enabled = False
    tmrTimer.Enabled = False

        '#### Remove the tray icon and move fraTimeRemaining back out of sight
    Call cTrayIcon.Delete
    fraTimeRemaining.Top = frmEggTimer.Height * 2
End Sub


'#########################################################
'# tmrTimer_Timer() incs/decs the global variables which runs the timer itself
'#########################################################
Private Sub tmrTimer_Timer()
        '#### If we are in 'timer' mode, simply increment the counters
    If (optType(2).Value) Then
            '#### Inc the lSeconds
        lSeconds = lSeconds + 1

            '#### If we have reached a full minute
        If (lSeconds = 60) Then
                '#### Reset lSeconds and inc lMinutes
            lSeconds = 0
            lMinutes = lMinutes + 1

                '#### If we have reached a full hour
            If (lMinutes = 60) Then
                    '#### Reset lMinutes and inc lHours
                lMinutes = 0
                lHours = lHours + 1
            End If
        End If

        '#### Else we are in one of the two countdown modes
    Else
            '#### If the time is up
        If (lHours = 0 And lMinutes = 0 And lSeconds = 0) Then
                '#### Disable the timer, remove the cTray Icon and beep the system's speaker
            tmrTimer.Enabled = False
            Call cTrayIcon.Delete
            Beep

                '#### Prompt the user, unload form and exit the sub in order to avoid the sTimeRemaining code below
            Call MsgBox("Ding! Times up!", vbOKOnly + vbInformation + vbSystemModal, "Cn's EggTimer")
            Unload Me
            Exit Sub

            '#### Else we need to decrement the time
        Else
                '#### If we've run out of seconds
            If (lSeconds = 0) Then
                    '#### Decrement lMinutes and reset lSeconds
                If (lMinutes > 0) Then lMinutes = lMinutes - 1
                lSeconds = 59

                '#### Else simply dec the lSeconds
            Else
                lSeconds = lSeconds - 1
            End If

                '#### If we've run out of lMinutes and lHours
            If (lMinutes = 0 And lHours > 0) Then
                    '#### Decrement lHours and reset lMinutes
                lHours = lHours - 1
                lMinutes = 59
            End If
        End If
    End If

        '#### Reset sTimeRemaining, zero padding if necessary
    sTimeRemaining = lHours
    If (lMinutes < 10) Then
        sTimeRemaining = sTimeRemaining & ":0" & lMinutes
    Else
        sTimeRemaining = sTimeRemaining & ":" & lMinutes
    End If
    If (lSeconds < 10) Then
        sTimeRemaining = sTimeRemaining & ":0" & lSeconds
    Else
        sTimeRemaining = sTimeRemaining & ":" & lSeconds
    End If

        '#### Update the cTrayIcon and lblTimeRemaining
    Call cTrayIcon.Modify(imgTimer, sTimeRemaining)
    lblTimeRemaining.Caption = sTimeRemaining
End Sub
