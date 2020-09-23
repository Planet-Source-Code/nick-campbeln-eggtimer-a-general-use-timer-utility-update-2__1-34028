VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#####################################################################################
'#  EggTimer - A General Use Timer Utility
'#      By: Nick Campbeln
'#
'#  Revision Histroy:
'#      1.1 (Apr 23, 2002):
'#          Cleaned up code/fixed some return value bugs
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



'#####################################################################################
'# Global declarations and assignments
'#####################################################################################
    '#### UDT that will be passed into Shell_NotifyIcon
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

    '#### Shell_NotifyIcon constants to add, modify, or delete an icon from the taskbar
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

    '#### Shell_NotifyIcon constants that indicate valid members of the NOTIFYICONDATA UDT
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

    '#### Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pNID As NOTIFYICONDATA) As Boolean

    '#### Constant for the message sent when a mouse event is generated within the boundries of the icon
Private Const WM_MOUSEMOVE = &H200

    '#### Setup all of the vars required by the class
Private oNID As NOTIFYICONDATA
Private sHoverText As String
Private bVisible As Boolean



'#####################################################################################
'# Class init/terminate subs
'#####################################################################################
Private Sub Class_Initialize()
    bVisible = False
End Sub
Private Sub Class_Terminate()
    Delete
End Sub



'#####################################################################################
'# Get/Put properties
'#####################################################################################
'#########################################################
'# Return the current sHoverText to the caller
'#########################################################
Public Property Get HoverText() As String
    HoverText = sHoverText
End Property



'#####################################################################################
'# Public subs/functions
'#####################################################################################
'#########################################################
'# Adds an icon into the system tray
'#########################################################
Public Function Add(ByRef frmForm As Form, ByRef imgImage As Image, sText As String) As Boolean
        '#### As long as the icon is not currently showing, add it
    If (Not bVisible) Then
            '#### Setup the UDT for the API call below
        With oNID
            .cbSize = Len(oNID)
            .hWnd = frmForm.hWnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = imgImage.Picture
            .szTip = sText & vbNullChar
        End With

            '#### Set the global vars, the return value and send the UDT at Shell_NotifyIcon to show the icon
        sHoverText = sText
        bVisible = True
        Add = True
        Call Shell_NotifyIcon(NIM_ADD, oNID)

        '#### Else there is already an icon for this class, so return false
    Else
        Add = False
    End If
End Function


'#########################################################
'# Updates the currently visible icon in the system tray with new icon and text
'#########################################################
Public Function Modify(ByRef imgImage As Image, sText As String) As Boolean
        '#### As long as the icon is currently showing, modify it
    If (bVisible) Then
            '#### Reset the values in the UDT
        With oNID
            .hIcon = imgImage.Picture
            .szTip = sText & vbNullChar
        End With

            '#### Reset sHoverText and send the UDT at Shell_NotifyIcon to update the icon
        sHoverText = sText
        Call Shell_NotifyIcon(NIM_MODIFY, oNID)
    End If

        '#### Return the value of bVisible to the caller
    Modify = bVisible
End Function


'#########################################################
'# Removes the currently visible icon in the system tray
'#########################################################
Public Function Delete() As Boolean
        '#### As long as the icon is currently showing, delete it
    If (bVisible) Then
            '#### Set bVisible, return value and remove the icon from the system tray
        bVisible = False
        Delete = True
        Call Shell_NotifyIcon(NIM_DELETE, oNID)

        '#### Else there was no icon to delete, so return false
    Else
        Delete = False
    End If
End Function

