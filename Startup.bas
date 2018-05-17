Attribute VB_Name = "Startup"
'---------------------------------------------------------------------
' Copyright © 2000-2002 SPACE.com Inc., New York, NY
'
' Permission is hereby granted to use this Software for any purpose
' including combining with commercial products, creating derivative
' works, and redistribution of source or binary code, without
' limitation or consideration. Any redistributed copies of this
' Software must include the above Copyright Notice.
'
' THIS SOFTWARE IS PROVIDED "AS IS". SPACE.COM, INC. MAKES NO
' WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
' SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'---------------------------------------------------------------------
'   ============
'   STARTUP.BAS
'   ============
'
' CyberDrive main startup module
'
' Written:  20-Jun-03   Jon Brewster
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 20-Jun-03 jab     Initial edit
' 09-Sept-04 MCH    Rip-off for Skybadger dome controller
' -----------------------------------------------------------------------------

Option Explicit

Public Const INVALID_COORDINATE As Double = -100000#
Public Const EMPTY_COORDINATE As Double = -50000#
Public Const DIR_CCW As Integer = 0
Public Const DIR_CW As Integer = 1
Public Const SLEW_MAX As Integer = 127
Public Const SLEW_MIN As Integer = 40

'------------------------
' Timer intervals (sec.)
'------------------------
Public Const TIMER_INTERVAL = 1                     ' sec per cycle

'-------------------
' ASCOM Identifiers
'-------------------
Public Const ID As String = "Skybadger.Dome"
Private Const DESC As String = "Skybadger Dome"
Private Const RegVer As String = "1.2"
Public Const ALERT_TITLE As String = "Skybadger"
Public Const INSTRUMENT_NAME As String = "Skybadger"
Public Const INSTRUMENT_DESCRIPTION As String = _
    "Skybadger Dome Driver"

'VB constants
Public Const vbLogEventTypeError = 1 'Error.
Public Const vbLogEventTypeWarning = 2 'Warning.
Public Const vbLogEventTypeInformation = 4 'Information.

' ----------
' Variables
' ----------
Public g_Profile As DriverHelper.Profile
Public g_Util As DriverHelper.Util
Public g_bRunExecutable As Boolean
Public g_iConnections As Integer

' ---------------
' State Variables
' ---------------
'movement control
Public g_dSlewSpeed As Double              ' degrees per sec
Public g_dStepSize As Double               ' degrees per GUI step
Public g_dGear As Double                   ' small gear revs per full turn
'Com port for motor and LCD interface
Public g_iSerPortID                        'ID of serial port used to talk to azimuth compass.
Public g_sSerPortSettings                  'setup string for the serial port for the compass
'I2C device addresses
Public g_iControllerAddress                ' I2C Address for dome motor controller
Public g_iLCDAddress                        'I2C Address of LCD display
Public g_iCompassAddress                   ' I2C address of encoder device
'Position
Public g_dDomeSyncOffset As Double         ' offset derived from calibration via sync
Public g_dDomeAz As Double                 ' Current Az for Dome
Public g_dSetPark As Double                ' Park position
Public g_dTargetAz As Double               ' Target Az
'state
Public g_bConnected As Boolean             ' Whether dome is connected
Public g_bAtPark As Boolean                ' Park state
Public g_eSlewing As Going                 ' Move in progress
Public g_bSlaved As Boolean                ' are we slaved


' ----------------------
' Other global variables
' ----------------------
Public g_handBox As frmHandBox             ' Hand box
Public g_show As frmShow                   ' Traffic window
Public g_timer As VB.timer                 ' Handy reference to timer
Public g_SerPort As MSCommLib.MSComm

'USB I2C connection
Public g_sDevSymName As String      ' symbolic name of USB device, example: "UsbI2cIo"
Public g_byDevInstance As Byte     ' currently selected device instance number
Public g_hDevInstance As Long       ' handle to the currently selected device instance
Dim g_bDevicePresent As Boolean     ' flag to indicate presence of device
'---------------------------------------------------------------------
'
' Main() - main entry point
'
'---------------------------------------------------------------------
Sub Main()
    'App.StartLogging App.Path & "skybadgerdome.log", CLng(&H32)
    'thread id with overwrite to file
        
    Set g_handBox = New frmHandBox
    Set g_timer = g_handBox.timer
    'Set g_SerPort = g_handBox.MSComm1
    Set g_SerPort = CreateObject("MSCommLib.MSComm")
    Set g_Profile = New DriverHelper.Profile
    Set g_Util = New DriverHelper.Util
    g_Profile.DeviceType = "Dome"                   ' We're a Dome driver
        
    LoadDLL "astro32.dll"                           ' Load the astronomy functions DLL
    LoadDLL "Vbhlp32.dll"                           ' Load the shifting functions
    LoadDLL "usbi2cio.dll"                          ' Load the Usb io functions
    g_sDevSymName = "UsbI2cIo"
    
    With App
        .StartLogging "C:\temp\skybadgerdome.log", vbLogToFile
        .LogEvent "Specified Path File Logging", 4
    End With
        
    If App.StartMode = vbSModeStandalone Then
        g_bRunExecutable = True                     ' launched via double click
        App.LogEvent "SkybadgerDome running stand-alone", vbLogEventTypeInformation
    Else
        g_bRunExecutable = False                    ' running as server only
        App.LogEvent "SkybadgerDome running as server", vbLogEventTypeInformation
    End If
    
    g_iConnections = 0                              ' zero connections currently
    g_bConnected = False                            ' Not yet connected
    g_bAtPark = True                                ' Assume start parked
    g_eSlewing = slewNowhere                        ' Not slewing
    g_dTargetAz = INVALID_COORDINATE                ' Set target=current
    g_dDomeAz = INVALID_COORDINATE
    
    g_Profile.Register ID, DESC                     ' Self reg (skips if already reg)
    
    
    ' Persistent settings - Create on first start
    If g_Profile.GetValue(ID, "RegVer") <> RegVer Then
        g_Profile.WriteValue ID, "RegVer", RegVer
        g_Profile.WriteValue ID, "SerPortSettings", "19200,n,8,1"
        g_Profile.WriteValue ID, "SerPortId", 1
        g_Profile.WriteValue ID, "Motor Controller Address", "&Hb0"
        g_Profile.WriteValue ID, "LCD Controller Address", "&Hc2"
        g_Profile.WriteValue ID, "Compass Address", "&Hc0"
        g_Profile.WriteValue ID, "SyncOffset", "0"
        g_Profile.WriteValue ID, "SetPark", "180"
        g_Profile.WriteValue ID, "SlewSpeed", "4"
        g_Profile.WriteValue ID, "StepSize", "5"
        g_Profile.WriteValue ID, "Gear", "84"
        g_Profile.WriteValue ID, "Left", "100"
        g_Profile.WriteValue ID, "Top", "100"
    '    g_Profile.WriteValue ID, "DomeCom", "3"
        App.LogEvent "SkybadgerDome registered profile ( none previous)", vbLogEventTypeInformation
    End If
    
    g_dSetPark = val(g_Profile.GetValue(ID, "SetPark"))
    g_dDomeSyncOffset = val(g_Profile.GetValue(ID, "SyncOffset"))
    g_dStepSize = val(g_Profile.GetValue(ID, "StepSize"))
    g_dGear = val(g_Profile.GetValue(ID, "Gear"))
    'i2c addresses
    g_iControllerAddress = val(g_Profile.GetValue(ID, "Motor Controller Address"))
    g_iCompassAddress = val(g_Profile.GetValue(ID, "Compass Address"))
    g_iLCDAddress = val(g_Profile.GetValue(ID, "LCD Controller Address"))
    'com port settings
    'g_iDomeCom = CInt(g_Profile.GetValue(ID, "DomeCom"))
    g_iSerPortID = val(g_Profile.GetValue(ID, "SerPortId"))
    g_sSerPortSettings = g_Profile.GetValue(ID, "SerPortSettings")

    'Setup serial port object
    On Error GoTo com_err
    'setSerial g_iSerPortID, False
    g_SerPort.Settings = g_sSerPortSettings
    g_SerPort.InputMode = 1 'Binary
    g_SerPort.InputLen = 0 ' all input at once
    g_SerPort.CommPort = g_iSerPortID
    GoTo no_comm_err
com_err:
    App.LogEvent "SkybadgerDome unable to contact dome via serial port ", vbLogEventTypeError
    MsgBox "Unable to setup comm port to dome"
no_comm_err:
    App.LogEvent "SkybadgerDome connected to dome via serial port ", vbLogEventTypeInformation
    Load frmSetup
    Load g_handBox
    
    With g_handBox
        .DomeAz = g_dDomeAz
        .LabelButtons
        .RefreshLEDs
        .Left = CLng(g_Profile.GetValue(ID, "Left")) * Screen.TwipsPerPixelX
        .Top = CLng(g_Profile.GetValue(ID, "Top")) * Screen.TwipsPerPixelY
        
        If .Left < 0 Then _
            .Left = 0
        If .Top < 0 Then _
            .Top = 0
    End With
         
    If g_bRunExecutable Then
        g_handBox.WindowState = vbNormal
    Else
        g_handBox.WindowState = vbMinimized
    End If
    
    g_handBox.Show
    g_timer.Interval = (TIMER_INTERVAL * 1000)  ' convert to millisec
    g_timer.Enabled = True
End Sub

'---------------------------------------------------------------------
'
' DoShutdown() - Handle handbox form Unload event
'
'---------------------------------------------------------------------
Sub DoShutdown()
    App.LogEvent "SkybadgerDome shutdown entered", vbLogEventTypeInformation
    g_timer.Enabled = False
    HW_Shutdown
    
    g_Profile.WriteValue ID, "SetPark", Str(g_dSetPark)
    g_Profile.WriteValue ID, "SyncOffset", Str(g_dDomeSyncOffset)
    g_Profile.WriteValue ID, "SlewSpeed", Str(g_dSlewSpeed)
    g_Profile.WriteValue ID, "StepSize", Str(g_dStepSize)
    g_Profile.WriteValue ID, "Park", Str(g_dSetPark)
    g_Profile.WriteValue ID, "Gear", Str(g_dGear)
    'I2C addresses
    g_Profile.WriteValue ID, "Motor Controller Address", Str(g_iControllerAddress)
    g_Profile.WriteValue ID, "Compass Address", Str(g_iCompassAddress)
    g_Profile.WriteValue ID, "LCD Controller Address", Str(g_iLCDAddress)
    'Com port settings
    'g_Profile.WriteValue ID, "DomeCom", Str(g_iDomeCom)
    g_Profile.WriteValue ID, "SerPortId", g_iSerPortID
    g_Profile.WriteValue ID, "SerPortSettings", g_sSerPortSettings

    g_handBox.Visible = True
    g_handBox.WindowState = vbNormal
    g_Profile.WriteValue ID, "Left", Str(g_handBox.Left \ Screen.TwipsPerPixelX)
    g_Profile.WriteValue ID, "Top", Str(g_handBox.Top \ Screen.TwipsPerPixelY)
    App.LogEvent "Skybadgerdome normal shutdown completed", vbLogEventTypeInformation
    
End Sub

'---------------------------------------------------------------------
'
' DoSetup() - Handle handbox Setup button click
'
'---------------------------------------------------------------------
Sub DoSetup()
    Dim ans As Boolean
    Dim newSerPortID
    App.LogEvent "SkybadgerDome entered DoSetup", vbLogEventTypeInformation
    
    With frmSetup
        .AllowUnload = False                        ' Assure not unloaded
        .SlewSpeed = g_dSlewSpeed
        .StepSize = g_dStepSize
        .Park = g_dSetPark
        .DomeCom = g_iSerPortID
    End With
    
    g_handBox.Visible = False                       ' May float over setup
    FloatWindow frmSetup.hwnd, True
    frmSetup.Show 1
    
    With frmSetup
        If .result Then             ' Unless cancelled
            g_dSlewSpeed = .SlewSpeed
            g_dStepSize = .StepSize
            g_dSetPark = .Park
            newSerPortID = .DomeCom
       End If
       
        .AllowUnload = True                     ' OK to unload now
    End With
    
    g_bAtPark = (g_dSetPark = g_dDomeAz)
    'Handle change of serial port.
       If newSerPortID <> g_iSerPortID Then
          setSerial newSerPortID, g_SerPort.PortOpen
          g_iSerPortID = newSerPortID
       End If
       
    With g_handBox
        .DomeAz = g_dDomeAz
        .LabelButtons
        .RefreshLEDs
    End With
    
    g_handBox.Visible = True
    
End Sub

'---------------------------------------------------------------------
'
' timer_tick() - Called by timer
'
'---------------------------------------------------------------------
Sub timer_tick()
    Dim button As Integer
    Dim dAz As Double
           
    'update current position - est. up to 0.5 secs delay for this.
    If (g_bConnected) Then
        If HW_GetAzimuth(dAz) Then g_dDomeAz = dAz
    End If
    
    '-----------------------
    ' Handle hand-box state
    '-----------------------
    button = g_handBox.ButtonState
    If g_bConnected And button <> 0 Then
        ' act on button
        Select Case (button)
            Case 1: ' Go anti-clockwise
                HW_Run DIR_CCW
                If Not g_show Is Nothing Then
                    If g_show.chkSlewing.Value = 1 Then _
                        g_show.TrafficLine "Slew CCW"
                End If
            Case 2: ' step clockwise
                HW_Move AzScale(g_dDomeAz - g_dStepSize)
                If Not g_show Is Nothing Then
                    If g_show.chkSlewing.Value = 1 Then _
                        g_show.TrafficLine "Step CCW"
                End If
            Case 3: ' Go counter clockwise
                HW_Run DIR_CW
                If Not g_show Is Nothing Then
                    If g_show.chkSlewing.Value = 1 Then _
                        g_show.TrafficLine "Slew CW"
                End If
            Case 4: ' step counter clockwise
                HW_Move AzScale(g_dDomeAz + g_dStepSize)
                If Not g_show Is Nothing Then
                    If g_show.chkSlewing.Value = 1 Then _
                        g_show.TrafficLine "Step CW"
                End If
            Case 5: ' EMERGENCY STOP
                HW_Halt
                If Not g_show Is Nothing Then
                    If g_show.chkSlewing.Value = 1 Then _
                        g_show.TrafficLine "Slew Halt"
                End If
            Case Else: ' other
                HW_Halt
                If Not g_show Is Nothing Then
                 If g_show.chkSlewing.Value = 1 Then _
                    g_show.TrafficLine "Unknown button pressed:" & button
                End If
        End Select
    End If

    '---------------------------------------------
    ' If we're slewing, see if we should stop yet
    '---------------------------------------------
    If g_bConnected Then
        If g_eSlewing <> slewNowhere Then
            ' If we're not just running in circles
            If g_eSlewing <> slewCW And g_eSlewing <> slewCCW Then
                ' are we there yet?
                If AzScale(g_dTargetAz - g_dDomeAz) < 2 Then
                    HW_Halt
                    If Not g_show Is Nothing Then
                        If g_show.chkSlewing.Value = 1 Then g_show.TrafficLine "(Slew completed)"
                    End If
                'slow down if we are getting nearer
                'ElseIf (AzScale(g_dTargetAz - g_dDomeAz) < 10) Then
                '    HW_Move g_dTargetAz
                Else
                    If Not g_show Is Nothing Then
                        If g_show.chkSlewing.Value = 1 Then _
                         g_show.TrafficLine "Slewing to : " & g_dTargetAz
                    End If
                End If
            End If
        Else
            ' check on motion anyway (temporary) ???
            If Not g_show Is Nothing Then
                If g_show.chkSlewing.Value = 1 Then HW_Slewing
            End If
        End If
    End If
    
    ' Update hand-box
    g_handBox.DomeAz = g_dDomeAz

    ' Update LCD display
    If (g_bConnected) Then
        HW_DisplayWriteTime
        HW_DisplayWriteAz
        HW_DisplayWriteState
    End If

End Sub

' ---------
' UTILITIES
' ---------
Public Function AzScale(Az As Double) As Double

    AzScale = Az Mod 360
    If AzScale < 0 Then AzScale = AzScale + 360#

End Function

Public Function FmtSexa(ByVal n As Double, ShowPlus As Boolean)
    Dim sg As String
    Dim us As String, ms As String, ss As String
    Dim u As Integer, m As Integer
    Dim fmt

    sg = "+"                                ' Assume positive
    If n < 0 Then                           ' Check neg.
        n = -n                              ' Make pos.
        sg = "-"                            ' Remember sign
    End If

    m = Fix(n)                              ' Units (deg or hr)
    us = Format$(m, "00")

    n = (n - m) * 60#
    m = Fix(n)                              ' Minutes
    ms = Format$(m, "00")

    n = (n - m) * 60#
    m = Fix(n)                              ' Minutes
    ss = Format$(m, "00")

    FmtSexa = us & ":" & ms & ":" & ss
    If ShowPlus Or (sg = "-") Then FmtSexa = sg & FmtSexa
    
End Function

Public Sub setSerial(ByVal portID As Integer, openPortForMe As Boolean)
'Setup serial port object
    If (g_SerPort.PortOpen = True) Then
        g_SerPort.PortOpen = False
    End If
    g_SerPort.Settings = g_sSerPortSettings
    g_SerPort.InputMode = 1 'Binary
    g_SerPort.InputLen = 0 ' all input at once
    g_SerPort.CommPort = g_iSerPortID
    If (openPortForMe = True) Then
        g_SerPort.PortOpen = True
    End If
End Sub
