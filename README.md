<div align="center">

## TaskManager


</div>

### Description

Here's a simple application to function like the Windows Task Manager...
 
### More Info
 
Start

'        a new project and add the following controls to the form:

'         Control   Name   Caption

'        

----

'        commandbutton cmdRefresh Refresh

'        commandbutton cmdSwitch  Switch

'        commandbutton cmdExit   Exit

'        listbox    lstApp


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-taskmanager__1-275/archive/master.zip)

### API Declarations

```
'for 16 bit (VB3 and VB4-16) use these:
  Declare Function ShowWindow Lib "User" _
            (ByVal hWnd As Integer,     ByVal flgs As Integer) _
            As Integer
        Declare Function GetWindow Lib "User" _
            (ByVal hWnd As Integer, ByVal wCmd As Integer) _
            As Integer
        Declare Function GetWindowWord Lib "User" _
            (ByVal hWnd As Integer, ByVal wIndx As Integer) _
            As Integer
        Declare Function GetWindowLong Lib "User" _
            (ByVal hWnd As Integer, ByVal wIndx As Integer) As Long
        Declare Function GetWindowText Lib "User" _
            (ByVal hWnd As Integer, ByVal lpSting As String, _
            ByVal nMaxCount As Integer) As Integer
        Declare Function GetWindowTextLength Lib "User" _
            (ByVal hWnd As Integer) As Integer
        Declare Function SetWindowPos Lib "User" _
            (ByVal hWnd As Integer, ByVal insaft As Integer, _
            ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, _
            ByVal flgs As Integer) As Integer
        Const WS_MINIMIZE = &H20000000 ' Style bit 'is minimized'
        Const HWND_TOP = 0       ' Move to top of z-order
        Const SWP_NOSIZE = &H1     ' Do not re-size window
        Const SWP_NOMOVE = &H2     ' Do not reposition window
        Const SWP_SHOWWINDOW = &H40   ' Make window visible/active
        Const GW_HWNDFIRST = 0     ' Get first Window handle
        Const GW_HWNDNEXT = 2      ' Get next window handle
        Const GWL_STYLE = (-16)     ' Get Window's style bits
        Const SW_RESTORE = 9      ' Restore window
        Dim IsTask As Long       ' Style bits for normal task
        ' The following bits will be combined to define properties
        ' of a 'normal' task top-level window. Any window with ' these set will be
        included in the list:
        Const WS_VISIBLE = &H10000000   ' Window is not hidden
        Const WS_BORDER = &H800000     ' Window has a border
        ' Other bits that are normally set include:
        Const WS_CLIPSIBLINGS = &H4000000 ' can clip windows
        Const WS_THICKFRAME = &H40000   ' Window has thick border
        Const WS_GROUP = &H20000      ' Window is top of group
        Const WS_TABSTOP = &H10000     ' Window has tabstop
 For VB4 32-bit change the function defintions to the following:
        Private Declare Function ShowWindow Lib "User32" _
            (ByVal hWnd As Long, ByVal flgs As Long) As Long
        Private Declare Function GetWindow Lib "User32" _
            (ByVal hWnd As Long, ByVal wCmd As Long) As Long
        Private Declare Function GetWindowWord Lib "User32" _
            (ByVal hWnd As Long, ByVal wIndx As Long) As Long
        Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" _
            (ByVal hWnd As Long, ByVal wIndx As Long) As Long
        Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" _
            (ByVal hWnd As Long, ByVal lpSting As String, ByVal nMaxCount As Long) As
        Long
        Private Declare Function GetWindowTextLength Lib "User32" Alias
        "GetWindowTextLengthA" _
            (ByVal hWnd As Long) As Long
        Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, _
            ByVal insaft As Long, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, _
            ByVal flgs As Long) As Long
```


### Source Code

```
Sub cmdExit_Click ()
        Unload Me        ' Get me out of here!
        Set activate = Nothing ' Kill Form reference for good measure
        End Sub
        Sub cmdRefresh_Click ()
        FindAllApps ' Update list of tasks
        End Sub
        Sub cmdSwitch_Click ()
        Dim hWnd As Long  ' handle to window
        Dim x As Long     ' work area
        Dim lngWW As Long   ' Window Style bits
        If lstApp.ListIndex < 0 Then Beep: Exit Sub
        ' Get window handle from listbox array
        hWnd = lstApp.ItemData(lstApp.ListIndex)
        ' Get style bits for window
        lngWW = GetWindowLong(hWnd, GWL_STYLE)
        ' If minimized do a restore
        If lngWW And WS_MINIMIZE Then
            x = ShowWindow(hWnd, SW_RESTORE)
        End If
        ' Move window to top of z-order/activate; no move/resize
        x = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
        End Sub
        Sub FindAllApps ()
        Dim hwCurr As Long
        Dim intLen As Long
        Dim strTitle As String
        ' process all top-level windows in master window list
        lstApp.Clear
        hwCurr = GetWindow(Me.hWnd, GW_HWNDFIRST) ' get first window
        Do While hwCurr ' repeat for all windows
         If hwCurr <> Me.hWnd And TaskWindow(hwCurr) Then
          intLen = GetWindowTextLength(hwCurr) + 1 ' Get length
          strTitle = Space$(intLen) ' Get caption
          intLen = GetWindowText(hwCurr, strTitle, intLen)
          If intLen > 0 Then ' If we have anything, add it
           lstApp.AddItem strTitle
        ' and let's save the window handle in the itemdata array
           lstApp.ItemData(lstApp.NewIndex) = hwCurr
          End If
         End If
         hwCurr = GetWindow(hwCurr, GW_HWNDNEXT)
        Loop
        End Sub
        Sub Form_Load ()
        IsTask = WS_VISIBLE Or WS_BORDER ' Define bits for normal task
        FindAllApps            ' Update list
        End Sub
        Sub Form_Paint ()
        FindAllApps ' Update List
        End Sub
        Sub Label1_Click ()
        FindAllApps ' Update list
        End Sub
        Sub lstApp_DblClick ()
        cmdSwitch.Value = True
        End Sub
        Function TaskWindow (hwCurr As Long) As Long
        Dim lngStyle As Long
        lngStyle = GetWindowLong(hwCurr, GWL_STYLE)
        If (lngStyle And IsTask) = IsTask Then TaskWindow = True
        End Function
```

