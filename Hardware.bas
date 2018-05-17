Attribute VB_Name = "Hardware"
' ==============
'  Hardware.BAS
' ==============
' CyberDrive hardware access module
' Written:  15-Jun-03   Jon Brewster
' Edits:
' When      Who     What
' --------- ---     -----------------------------------------------------------
' 15-Jun-03 jab     Initial edit
' 06 Sep 04 MCH     Ripped off for Skybadger
' -----------------------------------------------------------------------------
Private Const RowLength = 20
'Time row
Private Const iTimeValRow = 0
Private Const iTimeValOffset = 1
'Azimuth row
Private Const iAzimuthValRow = 1
Private Const iAzimuthValOffset = 1
'Dome slew state row
Private Const iStateValRow = 2
Private Const iStateValOffset = 1

Private Const sTimeTitle = "Time:"
Private Const sAzimuthTitle = "Az  :"
Private Const sStateTitle = "Dome:"

Option Explicit

Public Enum Going
    slewCCW = -1        ' just running till halt
    slewNowhere = 0     ' stopped, complete. not slewing
    slewCW = 1          ' just running till halt
    slewSomewhere = 2   ' specific Az based slew
    slewPark = 3        ' parking
    slewHome = 4        ' going home
End Enum

Private Const AccFactor = 16           ' acceleration in unknown units (old 0.8)
Private targetAz As Single               ' used for end of slew detection
Private speed As Double

Public Function OctetToHexStr(arrbytOctet)
 ' Function to convert OctetString (byte array) to Hex string.
 ' Code from Richard Mueller, a MS MVP in Scripting and ADSI

 Dim k
 OctetToHexStr = ""
 For k = 1 To LenB(arrbytOctet)
  OctetToHexStr = OctetToHexStr & Right("0" & Hex(AscB(MidB(arrbytOctet, k, 1))), 2)
 Next
 End Function

Public Function HW_Fetch() As Double
    Dim dAz As Double
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Fetch: "
    End If
    
    'Update if valid else ignore.
    If (HW_GetAzimuth(dAz)) Then
        g_dDomeAz = dAz
    End If
    HW_Fetch = g_dDomeAz
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd Format$(HW_Fetch, "000.0")
    End If
    
End Function
Public Sub HW_Halt()
    Dim Az As Single
    Dim dAz As Double
        
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Halt"
    End If

    HW_MotorsSetSpeedDirection 0, 1
    
    g_eSlewing = slewNowhere
    If HW_GetAzimuth(dAz) Then g_dDomeAz = dAz
    
    g_bAtPark = (g_dSetPark = g_dDomeAz)
    g_handBox.RefreshLEDs
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
    
End Sub

Public Sub HW_Init()
    Dim newAz As Single
    Dim dAz As Double
    Dim i As Integer
        
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Init"
    End If
    
    g_dDomeAz = g_dSetPark
    g_bAtPark = True
    g_handBox.RefreshLEDs
    g_eSlewing = slewNowhere
       
    ' connect to USB I2C port - controls motor
    For i = 0 To 127 Step 1
        g_hDevInstance = DAPI_OpenDeviceInstance(g_sDevSymName, i)
        If (g_hDevInstance <> INVALID_HANDLE_VALUE) Then
            g_byDevInstance = i
            ' succesfully opened a device
            Exit For
        End If
    Next i
    
    ' be careful with first command
    On Error GoTo CatchBeginDome
    'Controller setup commands here.
    HW_MotorsInit
    HW_MotorsSetSpeedDirection 0, DIR_CW
    
    GoTo FinalBeginDome
CatchBeginDome:
    Err.Raise SCODE_BAD_DOME, ERR_SOURCE, MSG_BAD_DOME
    Resume Next
FinalBeginDome:
    On Error GoTo 0
    
    On Error GoTo CatchGetEncoder
    'Connect to local serial port interface to radio I2C port
    If Not frmHandBox.MSComm1.PortOpen Then
      frmHandBox.MSComm1.PortOpen = True
    End If
    'Old code causes error due to bADLY REFERENCED SERIAL PORT OBJECT
    'If Not g_SerPort.PortOpen Then
    '    g_SerPort.PortOpen = True
    'End If
    If HW_GetAzimuth(dAz) Then
        g_dDomeAz = dAz
    End If
    
    On Error GoTo CatchGetDisplay
    'Can't write state till we have the compass bearing
    HW_DisplayInit
    HW_DisplayWriteTime
    HW_DisplayWriteAz
    HW_DisplayWriteState
    
    GoTo Final
CatchGetEncoder:
    Err.Raise SCODE_NO_DOME, ERR_SOURCE, MSG_NO_DOME
    Resume Next
    GoTo Final
CatchGetDisplay:
    Err.Raise SCODE_BAD_DISPLAY, ERR_SOURCE, MSG_BAD_DISPLAY
    Resume Next
    GoTo Final

Final:
    On Error GoTo 0
    
    g_dTargetAz = g_dDomeAz
    g_bConnected = True
    g_handBox.LabelButtons
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
    
End Sub

Public Sub HW_Move(Az As Double)
    Dim newEnc As Long

    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Move: " & Format$(Az, "000.0")
    End If
    
    g_bAtPark = False
    g_handBox.RefreshLEDs
    
    HW_MoveAzimuth CLng(Az Mod 360)
    g_dTargetAz = Az
    g_eSlewing = slewSomewhere
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd "Slew started, target: " & g_dTargetAz
    End If
    
End Sub

Public Sub HW_Park()
    Dim newEnc As Long
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Park"
    End If
    
    g_bAtPark = False
    g_handBox.RefreshLEDs
    
    g_dTargetAz = g_dSetPark
    HW_MoveAzimuth CLng(Fix(g_dTargetAz) Mod 360)

    g_eSlewing = slewPark
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd " (parking)"
    End If
    
End Sub

' Dir true means increasing 0 -> 180 -> 360 ie N thro E,
' as per sky rotation & compass coords
Public Sub HW_Run(Dir As Integer)
    Dim Enc As Long
    Dim speed As Single
    Dim direction As Single
        
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Run: " & IIf(Dir, "CW", "CCW")
    End If
    
    g_bAtPark = False
    g_handBox.RefreshLEDs
    
    'Add motor control commands in here.
    g_eSlewing = IIf(Dir = DIR_CW, slewCW, slewCCW)
    speed = 127 / 2
    direction = Dir
    HW_MotorsSetSpeedDirection speed, direction
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd " (running)"
    End If
    
End Sub

Public Sub HW_Shutdown()
   
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Shutdown"
    End If
    
    g_eSlewing = slewNowhere
    g_dDomeAz = INVALID_COORDINATE
    g_bAtPark = False
    g_handBox.RefreshLEDs
    g_bConnected = False
    
    If frmHandBox.MSComm1.PortOpen Then
        ' clean and close the port
       HW_DisplayInit
       frmHandBox.MSComm1.PortOpen = False
    End If
    
    g_handBox.LabelButtons
    g_handBox.DomeAz = g_dDomeAz
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
       
    If g_hDevInstance <> INVALID_HANDLE_VALUE Then
        HW_MotorsInit
        If DAPI_CloseDeviceInstance(g_hDevInstance) Then
      ' everythings zen
        If Not g_show Is Nothing Then
           If g_show.chkHW.Value = 1 Then _
              g_show.TrafficEnd "I2C Motor controller closed"
        End If
    Else
      ' SNAFU
      If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd "I2C Motor controller failed to close"
      End If
    End If
    g_hDevInstance = INVALID_HANDLE_VALUE
  End If
    
End Sub
Public Function HW_Slewing() As Boolean
    HW_Slewing = False
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Slewing: "
    End If
    
    If g_eSlewing = slewSomewhere Then
        '        If g_show.chkSlewing.Value = 1 Then _
        '            g_show.TrafficLine "(Slew aborted)"
        '    End If
        HW_Slewing = True
    Else
        HW_Slewing = False
    End If
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd CStr(HW_Slewing)
    End If
    
End Function

Public Sub HW_Sync(Az As Double)
    Dim dAz As Double
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficStart "HW_Sync: " & Format$(Az, "000.0")
    End If
    
    g_eSlewing = slewNowhere
    'set it to zero to get true reading of azimuth
    g_dDomeSyncOffset = 0#
    If (HW_GetAzimuth(dAz)) Then g_dDomeAz = dAz
    g_dDomeSyncOffset = Az - g_dDomeAz

    'update current to synced reading
    g_dDomeAz = CDbl(Az)
    g_bAtPark = (g_dSetPark = g_dDomeAz)
    g_handBox.RefreshLEDs
    
    If Not g_show Is Nothing Then
        If g_show.chkHW.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
    
End Sub


'===========================================
' Internal HW specific routines start here
'===========================================
Private Sub HW_MoveAzimuth(newVal As Long)
    Dim oldVal As Single
    Dim angle
    Dim direction
    Dim speed
    
    If Not g_show Is Nothing Then
        If g_show.chkEnc.Value = 1 Then _
            g_show.TrafficStart "MoveAzimuth: " & newVal
    End If
    
    oldVal = CSng(Fix(g_dDomeAz))
    angle = newVal - oldVal
    'Angle postive means CW
    direction = IIf(angle > 0, DIR_CW, DIR_CCW)
    
    'Speed is proportional to distance within limits, unless set from panel, within limits.
    If (g_dSlewSpeed <> 0#) Then
        speed = Fix(g_dSlewSpeed)
    Else
        speed = angle * 128 / 180
    End If
    speed = IIf(speed < SLEW_MIN, SLEW_MIN, speed)
    speed = IIf(speed > SLEW_MAX, SLEW_MAX, speed)
    
    If angle > 180 Then 'swap rotation size and direction
        angle = 360 - (newVal - oldVal)
        direction = IIf(direction = DIR_CW, DIR_CCW, DIR_CW)
        HW_MotorsSetSpeedDirection speed, direction
    ElseIf angle < -180 Then 'swap rotation size and direction
        angle = 360 + angle
        direction = IIf(direction = DIR_CW, DIR_CCW, DIR_CW)
        HW_MotorsSetSpeedDirection speed, direction
    ElseIf Abs(angle) >= 2 Then 'take what we have got
        HW_MotorsSetSpeedDirection speed, direction
    Else
        'nothing to do - increment too small.
    End If
    
    If Not g_show Is Nothing Then
        If g_show.chkEnc.Value = 1 Then _
            g_show.TrafficEnd " (moving enc)"
    End If

End Sub
Private Sub HW_DisplaySetCursorPosition(pos As Integer)
Dim Data As I2C_TRANS
Dim ret As Long
    Data.wCount.lo = 2
    Data.wCount.hi = 0
    Data.Data(0) = CByte(2)  'CMD to set cursor offset from origin,
    Data.Data(1) = CByte(pos)  'character position row 0 = 0-19, row1 = 20-39 etc
    Data.wMemAddr.lo = 0
    Data.wMemAddr.hi = 0
    Data.byType = I2C_TRANS_8ADR
    Data.byDevId = g_iLCDAddress
    ret = HW_I2CWrite(Data)
End Sub
Private Sub HW_DisplayInit()
Dim Data As I2C_TRANS
Dim i As Integer
Dim ret As Long
    Data.wCount.lo = 1
    Data.wCount.hi = 0
    Data.Data(0) = 12 'clearscreen
    Data.wMemAddr.lo = 0
    Data.wMemAddr.hi = 0
    Data.byType = I2C_TRANS_8ADR
    Data.byDevId = g_iLCDAddress
    ret = HW_I2CWrite(Data)
    
    Data.wCount.lo = 1
    Data.wCount.hi = 0
    Data.Data(0) = 5 'underline cursor
    Data.wMemAddr.lo = 0
    Data.wMemAddr.hi = 0
    Data.byType = I2C_TRANS_8ADR
    Data.byDevId = g_iLCDAddress
    ret = HW_I2CWrite(Data)
    
    Data.wCount.lo = 1
    Data.wCount.hi = 0
    Data.Data(0) = CByte(1) 'Home cursor
    Data.wMemAddr.lo = 0
    Data.wMemAddr.hi = 0
    Data.byType = I2C_TRANS_8ADR
    Data.byDevId = g_iLCDAddress
    ret = HW_I2CWrite(Data)
End Sub

Private Sub HW_DisplayWriteString(output As String)
Dim Data As I2C_TRANS
Dim i As Integer
Dim ret As Long
    Data.wCount.lo = Len(output)
    Data.wCount.hi = 0
    For i = 1 To Len(output)
        Data.Data(i - 1) = CByte(Asc(Mid(output, i, 1)))
    Next
    Data.wMemAddr.lo = 0
    Data.wMemAddr.hi = 0
    Data.byType = I2C_TRANS_8ADR
    Data.byDevId = g_iLCDAddress
    ret = HW_I2CWrite(Data)
End Sub
Public Sub HW_DisplayWriteTime()
   Dim timeNow As Date
   Dim timeString As String
   timeNow = Time
   timeString = sTimeTitle & Format(timeNow, " hh:mm:ss")
   HW_DisplaySetCursorPosition iTimeValOffset + iTimeValRow * RowLength
   HW_DisplayWriteString timeString
End Sub
Public Sub HW_DisplayWriteAz()
   Dim azString As String
   azString = sAzimuthTitle & Format(g_dDomeAz, " 000.0")
   HW_DisplaySetCursorPosition iAzimuthValOffset + (iAzimuthValRow * RowLength)
   HW_DisplayWriteString azString
End Sub
Public Sub HW_DisplayWriteState()
   Dim stateString As String
   Select Case g_eSlewing
   Case slewCCW:         ' just running till halt
        stateString = sStateTitle & " Slew CCW"
   Case slewNowhere:     ' stopped, complete. not slewing
        stateString = sStateTitle & " Halted  "
   Case slewCW:          ' just running till halt
        stateString = sStateTitle & " Slew CW "
   Case slewSomewhere:   ' specific Az based slew
        stateString = sStateTitle & " Tracking"
   Case slewPark:        ' parking
        stateString = sStateTitle & " Parking "
   Case slewHome:        ' going home
        stateString = sStateTitle & " Homing  "
   Case Else
        stateString = sStateTitle & " Unknown "
   End Select
   HW_DisplaySetCursorPosition iStateValOffset + iStateValRow * RowLength
   HW_DisplayWriteString stateString
End Sub
'Function to draw strings to I2C LCD display on motor controller
Private Function HW_DisplayWrite(ByRef firstString As String, ByRef secondString As String)
Dim Data As I2C_TRANS
Dim ret As Long
Dim i As Integer
    Data.wCount.lo = Len(firstString) Mod 64# 'must be less than 64 bytes
    Data.wCount.hi = 0
    For i = 0 To Len(firstString)
        Data.Data(i) = Chr(Mid(firstString, i, 1))
    Next
    Data.wMemAddr.lo = 0
    Data.wMemAddr.hi = 0
    Data.byType = I2C_TRANS_8ADR
    Data.byDevId = g_iLCDAddress
    ret = HW_I2CWrite(Data)
    
    Data.wCount.lo = Len(secondString) Mod 64#
    Data.wCount.hi = 0
    For i = 0 To Len(secondString)
        Data.Data(i) = Chr(Mid(secondString, i, 1))
    Next
    
    Data.wMemAddr.lo = 0
    Data.wMemAddr.hi = 0
    Data.byType = I2C_TRANS_8ADR
    Data.byDevId = g_iLCDAddress
    ret = HW_I2CWrite(Data)
    
    HW_DisplayWrite = ret
End Function
Public Function HW_GetAzimuth(ByRef result As Double) As Boolean
    Dim i As Integer
    Dim out As String
    Dim rtn As Double
    Dim hexstring As String

    Dim buffer() As Byte
    
    'Set up serial string to read compass
    out = Chr(&H55) & Chr(g_iCompassAddress Or &H1) & Chr(&H2) & Chr(&H2)
    'String to test using internal battery
    'out = Chr(&H5A) & Chr(3 Or &H1) & Chr(&O0) & Chr(&O0)
    
    If Not g_show Is Nothing Then
        If g_show.chkBytes.Value = 1 Then
            g_show.TrafficStart "Sending " & ":"
            For i = 1 To Len(out)
                g_show.TrafficChar Hex(Asc(Mid(out, i, 1)))
            Next i
        End If
    End If
    
    frmHandBox.MSComm1.output = out
    
    'add a serial poll timer and check for two bytes returned - the high byte and low byte
    Dim timedOut As Boolean
    timedOut = False
    Dim startTime, endTime
    startTime = timer()
    Do While (frmHandBox.MSComm1.InBufferCount < 2) And (timer() - startTime) < 0.5
        'If Not g_show Is Nothing Then
        '    If g_show.chkEnc.Value = 1 Then _
        '        g_show.TrafficLine "azimuth polling delay "
        'End If
    Loop
    
    'Currently doesn't handle midnight wrap-around
    If (timer > startTime + 0.5) Then
        timedOut = True
        g_handBox.ErrorLED
        result = -1
        HW_GetAzimuth = False
    ElseIf frmHandBox.MSComm1.InBufferCount >= 2 Then
        buffer = frmHandBox.MSComm1.Input
        'rtn = (buffer(1) * 256) + buffer(0) / 10#
        hexstring = OctetToHexStr(buffer)
        rtn = CDbl("&H" & hexstring) / 10#
        If Not g_show Is Nothing Then
            If g_show.chkEnc.Value = 1 Then _
                g_show.TrafficLine "compass value: " & CStr(rtn)
        End If
        If (rtn >= 0# And rtn <= 360#) Then
            result = AzScale(g_dDomeSyncOffset + rtn)
            HW_GetAzimuth = True
        Else
            HW_GetAzimuth = False
        End If
    Else
        HW_GetAzimuth = False
    End If
    
End Function

Private Function Hex(char As Integer) As String
    Dim nib As Integer
    
    'Hex = CStr(char)
    'Exit Function
    
    nib = char \ 16
    If nib > 9 Then _
       nib = nib + 7
    Hex = "0x" & Chr(nib + Asc("0"))
    nib = char And &HF
    If nib > 9 Then _
       nib = nib + 7
    Hex = Hex & Chr(nib + Asc("0"))
    
End Function

Private Sub HW_MotorsInit()
Dim transaction As I2C_TRANS
Dim result As Boolean
    transaction.wMemAddr.lo = 0
    transaction.wMemAddr.hi = 0
    transaction.wCount.lo = 4
    transaction.wCount.hi = 0
    transaction.Data(0) = &H1  'motor control mode : two motors, -128 reverse, 0 stop, 127 full ahead
    transaction.Data(1) = &H0  'motor 1 is off
    transaction.Data(2) = &H0  'motor 2 is off
    transaction.Data(2) = &H80 'acceleration is 16 ms per step
    transaction.byType = I2C_TRANS_8ADR
    transaction.byDevId = g_iControllerAddress
    result = HW_I2CWrite(transaction)
End Sub
' 1 is increasing bearing, -1 is decreasing
Private Sub HW_MotorsSetSpeedDirection(ByVal size As Integer, ByVal direction As Integer)
Dim transaction As I2C_TRANS

size = Abs(size) Mod 127

If direction = DIR_CCW Then
    'transaction.Data(0) = (&H80 Or size )
    size = 255 - size
End If

transaction.Data(0) = size
transaction.Data(1) = size

If Not g_show Is Nothing Then
        If g_show.chkBytes.Value = 1 Then g_show.TrafficLine "Motor speed: " & size & ", dirn: " & direction
End If

    transaction.wMemAddr.lo = 1
    transaction.wMemAddr.hi = 0
    transaction.wCount.lo = 2
    transaction.wCount.hi = 0
    transaction.byType = I2C_TRANS_8ADR
    transaction.byDevId = g_iControllerAddress
    Dim success As Integer
    success = HW_I2CWrite(transaction)
End Sub

Private Function HW_I2CWrite(Data As I2C_TRANS) As Boolean
  ' Perform an I2C Write transaction to an 8 bit I2C device
  
  Dim lWritten As Long                          ' Dimension a long to hold the returned value
  
  ' call the function
  lWritten = DAPI_WriteI2c(g_hDevInstance, Data)

  lWritten = Data.wCount.lo + Data.wCount.hi * 256
  If (lWritten) Then
    ' function call ok
    HW_I2CWrite = True
  Else
    ' function call failed
    HW_I2CWrite = False
    'Call MsgBox("Incorrect Return value", vbOKOnly, " Error calling DAPI_WriteI2C() function")
  End If
  
End Function

Public Function HW_I2CRead(Data As I2C_TRANS) As Boolean
  ' Perform an I2C Read transaction
  Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
  Dim lRead As Long                          ' Dimension a long to hold the returned value
  
  ' call the function
  lRead = DAPI_ReadI2c(g_hDevInstance, Data)

  If (lRead = Data.wCount.lo + (Data.wCount.hi * 256)) Then
    ' function call ok
    HW_I2CRead = True
  Else
    ' function call failed
    HW_I2CRead = False
    'Call MsgBox("Incorrect Return value", vbOKOnly, " Error calling DAPI_ReadI2C() function")
  End If

End Function
