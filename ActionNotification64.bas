Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright Â©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.


'defWindowProc holds the address
'of the default window message processing
'procedure returned by SetWindowLong
Public defWindowProc As LongPtr

'flag preventing re-creating the timer
Private tmrRunning As Boolean

'Get/SetWindowLong messages
Const GWL_WNDPROC = (-4)
Const GWL_HINSTANCE = (-6)
Const GWL_HWNDPARENT = (-8)
Const GWL_STYLE = (-16)
Const GWL_EXSTYLE = (-20)
Const GWL_USERDATA = (-21)
Const GWL_ID = (-12)

'general windows messages
Private Const WM_USER As Long = &H400
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10
Private Const WM_TIMER = &H113

'mouse constants for the callback
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206

'WM_MYHOOK is a private message the shell_notify api
'will pass to WindowProc when the systray icon is acted upon
Private Const WM_APP As Long = &H8000&
Public Const WM_MYHOOK As Long = WM_APP + &H15

'ID constant representing this
'application in the systray. Use
'a unique ID for each systray icon
'your app will add in order to
'differentiate between icons selected.
'The ID is returned in wParam.
Private Const APP_SYSTRAY_ID = 999

'ID constant representing this
'application for SetTimer
Public Const APP_TIMER_EVENT_ID As Long = 998

'const holding number of milliseconds to timeout
'7000=7 seconds
Public Const APP_TIMER_MILLISECONDS As Long = 7000

'balloon tip notification messages
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private Const NOTIFYICONDATA_SIZE As Long = NOTIFYICONDATA_V3_SIZE


Private Const NOTIFYICON_VERSION = &H3

'shell_notify flags
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
'shell_notify messages
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5
'shell_notify styles
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

'shell_notify icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As LongPtr
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As LongPtr
        szTip As String * 128
        dwState As Long
        dwStateMask As Long
        szInfo As String * 256
        uTimeoutAndVersion As Long
        szInfoTitle As String * 64
        dwInfoFlags As Long
        guidItem As GUID
'        hBalloonIcon As LongPtr
End Type


' http://www.cadsharp.com/docs/Win32API_PtrSafe.txt

Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long


#If Win64 Then
Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr

Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Private Declare PtrSafe Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
      

      
Public Sub ShellTrayIconAdd(hwnd As LongPtr, _
                            hIcon As Long, _
                            sToolTip As String)
   
   Dim nid As NOTIFYICONDATA
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP Or NIF_INFO
      .dwState = NIS_SHAREDICON
      .hIcon = hIcon
      .szTip = sToolTip & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      .uCallbackMessage = WM_MYHOOK
   End With
   
  'add the icon ...
   If Shell_NotifyIcon(NIM_ADD, nid) = 1 Then
   
     '... and inform the system of the
     'NOTIFYICON version in use
      Call Shell_NotifyIcon(NIM_SETVERSION, nid)
      
     'prepare to receive the systray messages
      Call SubClass(hwnd)
      
   End If
       
End Sub


Public Sub ShellTrayIconRemove(hwnd As LongPtr)

   Dim nid As NOTIFYICONDATA
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
   End With
   
   If tmrRunning Then Call TimerStop(hwnd)
   Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub


Private Sub ShellTrayBalloonTipClose(hwnd As LongPtr)

   Dim nid As NOTIFYICONDATA
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_TIP Or NIF_INFO
      .szTip = vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
   End With
   
   Call Shell_NotifyIcon(NIM_MODIFY, nid)
   
End Sub


Public Sub ShellTrayBalloonTipShow(hwnd As LongPtr, _
                                   nIconIndex As Long, _
                                   sTitle As String, _
                                   sMessage As String)

   Dim nid As NOTIFYICONDATA
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      .szInfoTitle = sTitle & vbNullChar
      .szInfo = sMessage & vbNullChar
   End With

   Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub


Private Sub SubClass(hwnd As LongPtr)

  'assign our own window message
  'procedure (WindowProc)
   On Error Resume Next
   defWindowProc = SetWindowLongPtr(hwnd, GWL_WNDPROC, AddressOf WindowProc)
   
End Sub


Public Sub UnSubClass(hwnd As LongPtr)

  'restore the default message handling
  'before exiting
   If defWindowProc <> 0 Then
      SetWindowLongPtr hwnd, GWL_WNDPROC, defWindowProc
      defWindowProc = 0
   End If
   
End Sub


Private Sub TimerBegin(ByVal hwndOwner As LongPtr, ByVal dwMilliseconds As Long)

   If Not tmrRunning Then

      If dwMilliseconds <> 0 Then

        'SetTimer returns the event ID we
        'assign if it starts successfully,
        'so this is assigned to the Boolean
        'flag to indicate the timer is running.
         tmrRunning = SetTimer(hwndOwner, _
                               APP_TIMER_EVENT_ID, _
                               dwMilliseconds, _
                               AddressOf TimerProc) = APP_TIMER_EVENT_ID
         
         Debug.Print "timer started"

      End If

   End If

End Sub


Public Function TimerProc(ByVal hwnd As LongPtr, _
                          ByVal uMsg As Long, _
                          ByVal idEvent As Long, _
                          ByVal dwTime As Long) As Long

   Select Case uMsg
      Case WM_TIMER

         If idEvent = APP_TIMER_EVENT_ID Then
            If tmrRunning = True Then
            
               Debug.Print "timer proc fired"
               Debug.Print "  shutting down balloon"
               Call TimerStop(hwnd)
               Call ShellTrayBalloonTipClose(Form1.hwnd)
               
            End If  'tmrRunning
         End If  'idEvent

      Case Else
   End Select

End Function


Private Sub TimerStop(ByVal hwnd As LongPtr)

   If tmrRunning = True Then

      Debug.Print "timer stopped"
      Call KillTimer(hwnd, APP_TIMER_EVENT_ID)
      tmrRunning = False

   End If

End Sub


Public Function WindowProc(ByVal hwnd As LongPtr, _
                           ByVal uMsg As Long, _
                           ByVal wParam As LongPtr, _
                           ByVal lParam As LongPtr) As Long

  'If the handle returned is to our form,
  'call a message handler to deal with
  'tray notifications. If it is a general
  'system message, pass it on to
  'the default window procedure.
  '
  'If destined for the form and equal to
  'our custom hook message (WM_MYHOOK),
  'examining lParam reveals the message
  'generated, to which we react appropriately.
   On Error Resume Next
  
   Select Case hwnd
   
     'form-specific handler
      Case Form1.hwnd
         
         Select Case uMsg
          'check uMsg for the application-defined
          'identifier (NID.uID) assigned to the
          'systray icon in NOTIFYICONDATA (NID).
  
           'WM_MYHOOK was defined as the message sent
           'as the .uCallbackMessage member of
           'NOTIFYICONDATA the systray icon
            Case WM_MYHOOK
            
              'lParam is the value of the message
              'that generated the tray notification.
               Select Case lParam
                  Case WM_RBUTTONUP

                 'This assures that focus is restored to
                 'the form when the menu is closed. If the
                 'form is hidden, it (correctly) has no effect.
                  Call SetForegroundWindow(Form1.hwnd)
                  Form1.PopupMenu Form1.zmnuSysTrayDemo
               
                  Case NIN_BALLOONSHOW
                    'the balloon tip has just appeared so
                    'set the timer to automatically close it
                     Call TimerBegin(hwnd, APP_TIMER_MILLISECONDS)
                     Debug.Print "NIN_BALLOONSHOW"
     
                  Case NIN_BALLOONHIDE
                    'the balloon tip has just been hidden,
                    'either because of a user-click, the
                    'system timeout being reached, or our
                    'SetTimer timeout expiring, so ensure
                    'the timer has stopped.
                     Call TimerStop(hwnd)
                     Debug.Print "NIN_BALLOONHIDE"

                  Case NIN_BALLOONUSERCLICK
                    'the balloon tip was clicked so
                    'ensure the timer won't fire
                     Call TimerStop(hwnd)
                     Debug.Print "NIN_BALLOONUSERCLICK"
                            
                  Case NIN_BALLOONTIMEOUT
                    'the system timeout has been reached
                    'which causes the system to close the
                    'tip without intervention. The timer
                    'must also be stopped now. Note that
                    'this message does not fire if the
                    'balloon tip is closed through our
                    'SetTimer method!
                     Call TimerStop(hwnd)
                     Debug.Print "NIN_BALLOONTIMEOUT"
               
               End Select
            
           'handle any other form messages by
           'passing to the default message proc
            Case Else
            
               WindowProc = CallWindowProc(defWindowProc, _
                                            hwnd, _
                                            uMsg, _
                                            wParam, _
                                            lParam)
               Exit Function
            
         End Select
     
     'this takes care of messages when the
     'handle specified is not that of the form
      Case Else
      
          WindowProc = CallWindowProc(defWindowProc, _
                                      hwnd, _
                                      uMsg, _
                                      wParam, _
                                      lParam)
   End Select
   
End Function


Public Function CustomStatus(ByVal nIconIndex As Long, _
                             ByVal sTitle As String, _
                             ByVal sMessage As String)

    Call ShellTrayIconAdd(0, 0, "Custom Outlook Status")
    

    Call ShellTrayBalloonTipShow(0, nIconIndex, sTitle, sMessage)
 
End Function

