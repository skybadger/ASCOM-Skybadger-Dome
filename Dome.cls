VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dome"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ASCOM driver for Cyberdrive Dome"
'   =============
'   DOME.CLS
'   =============
'
' Implementation of the ASCOM Dome Class
'
' Written:  20-Jun-03   Jon Brewster
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 20-Jun-03 jab     Initial edit
'---------------------------------------------------------------------

Option Explicit

' --------------------
' Enumerated Constants
' --------------------
'
' These appear as enumerated constants to COM/ActiveX clients
'
Public Enum ShutterState
    shutterOpen = 0
    shutterClosed = 1
    shutterOpening = 2
    shutterClosing = 3
    shutterError = 4
End Enum

' ------------
' Class Events
' ------------

Private Sub Class_Initialize()
    
    If Not g_bRunExecutable Then _
        g_iConnections = g_iConnections + 1
      
End Sub

Private Sub Class_Terminate()
    
    If Not g_bRunExecutable Then
        If g_iConnections > 1 Then
            g_iConnections = g_iConnections - 1
        Else
            Unload g_handBox
        End If
    End If
    
End Sub

' ----------
' Properties
' ----------

Public Property Get Altitude() As Double

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "Altitude: "
    End If

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
        "Property Altitude" & MSG_NOT_IMPLEMENTED
    
End Property

Public Property Get AtHome() As Boolean

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "AtHome: "
    End If
    
    check_connected
    'Don't need one for skybadger dome
    AtHome = False

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd CStr(AtHome)
    End If
    
End Property

Public Property Get AtPark() As Boolean

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "AtPark: "
    End If
    
    check_connected
    AtPark = g_bAtPark

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd CStr(AtPark)
    End If
    
End Property

Public Property Get Azimuth() As Double

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "Azimuth: "
    End If
 
    check_connected
    Azimuth = g_dDomeAz

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd FmtSexa(Azimuth, False)
    End If
    
End Property

Public Property Get CanFindHome() As Boolean

    CanFindHome = True

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "CanFindHome: " & CanFindHome
    End If
    
End Property

Public Property Get CanPark() As Boolean

    CanPark = True
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "CanPark: " & CanPark
    End If
    
End Property

Public Property Get CanSetAltitude() As Boolean

    CanSetAltitude = False
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "CanSetAltitude: " & CanSetAltitude
    End If
    
End Property

Public Property Get CanSetAzimuth() As Boolean

    CanSetAzimuth = True
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "CanSetAzimuth: " & CanSetAzimuth
    End If
    
End Property

Public Property Get CanSetPark() As Boolean

    CanSetPark = True
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "CanSetPark: " & CanSetPark
    End If
    
End Property

Public Property Get CanSetShutter() As Boolean

    CanSetShutter = False

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "CanSetShutter: " & CanSetShutter
    End If
    
End Property

Public Property Get CanSlave() As Boolean

    CanSlave = True
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "CanSlave: " & CanSlave
    End If
    
End Property

Public Property Get CanSyncAzimuth() As Boolean

    CanSyncAzimuth = True
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "CanSyncAzimuth: " & CanSyncAzimuth
    End If
    
End Property

Public Property Get Connected() As Boolean

    Connected = g_bConnected
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "Connected: " & Connected
    End If
    
End Property

Public Property Let Connected(ByVal newVal As Boolean)
    
    Dim out As String

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "Connected: " & g_bConnected & " -> " & newVal
    End If
    
    out = " (done)"
    
    If g_bConnected <> newVal Then
        If newVal Then
            HW_Init
        Else
            If g_bRunExecutable Then
                out = " (ignored, not auto launch)"
            Else
                If g_iConnections < 2 Then
                    HW_Shutdown
                Else
                    out = " (ignored, not last user)"
                End If
            End If
        End If
    End If
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd out
    End If
    
End Property

Public Property Get Description() As String

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "Description"
    End If
    
    Description = INSTRUMENT_DESCRIPTION
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
    
End Property

Public Property Get DriverInfo() As String

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "DriverInfo"
    End If
    
    '
    ' Use the Project/Properties sheet, Make tab, to set these
    ' items. That way they will show in the Version tab of the
    ' Explorer property sheet, and the exact same data will
    ' show in Dome.DriverInfo.
    '
    DriverInfo = App.FileDescription & " " & _
                App.Major & "." & App.Minor & "." & App.Revision
    If App.CompanyName <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.CompanyName
    If App.LegalCopyright <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.LegalCopyright
    If App.Comments <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.Comments
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
        
End Property

Public Property Get InterfaceVersion() As Integer

    InterfaceVersion = 1
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "InterfaceVersion: " & InterfaceVersion
    End If
    
End Property

Public Property Get Name() As String

    Name = INSTRUMENT_NAME
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "Name: " & Name
    End If
    
End Property

Public Property Get ShutterStatus() As ShutterState

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "ShutterStatus: "
    End If
    
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
        "Property ShutterStatus" & MSG_NOT_IMPLEMENTED
        
End Property

Public Property Get Slaved() As Boolean

    Slaved = g_bSlaved
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "Slaved: " & Slaved
    End If
    
End Property

Public Property Let Slaved(ByVal newVal As Boolean)
        
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficLine "Slaved: -> " & newVal
    End If
    
    g_bSlaved = newVal
    'If newVal Then _
        Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
            "Property Slaved" & MSG_NOT_IMPLEMENTED
        
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
    
End Property

Public Property Get Slewing() As Boolean

    Dim out As String
    
    If Not g_show Is Nothing Then
        If g_show.chkSlewing.Value = 1 Then
            Select Case g_eSlewing
                Case slewNowhere:   out = "Not Slewing"
                Case slewHome:      out = "Homing"
                Case slewPark:      out = "Parking"
                Case slewSomewhere: out = "Slewing"
                Case slewCW:        out = "Running CW"
                Case slewCCW:       out = "Running CCW"
            End Select
            g_show.TrafficChar out
            End If
    End If
    
    check_connected
    Slewing = (g_eSlewing <> slewNowhere)
    
End Property

' -------
' Methods
' -------

Public Sub AbortSlew()

    If Not g_show Is Nothing Then
        If g_show.chkSlew.Value = 1 Then _
            g_show.TrafficStart "AbortSlew"
    End If
    
    check_connected
    HW_Halt
    
    If Not g_show Is Nothing Then
        If g_show.chkSlew.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
    
End Sub

Public Sub CloseShutter()

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "CloseShutter"
    End If
    
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
        "Method CloseShutter" & MSG_NOT_IMPLEMENTED
    
End Sub

' should implement ???
Public Sub CommandBlind(ByVal Command As String)

        Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
            "Method CommandBlind" & MSG_NOT_IMPLEMENTED
            
End Sub

Public Function CommandBool(ByVal Command As String) As Boolean

        Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
            "Method CommandBool" & MSG_NOT_IMPLEMENTED
            
End Function

' should implement ???
Public Function CommandString(ByVal Command As String) As String

        'Maybe use this to provide access to I2c devices
        
        Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
            "Method CommandString" & MSG_NOT_IMPLEMENTED
            
End Function

Public Sub FindHome()

    If Not g_show Is Nothing Then
        If g_show.chkSlew.Value = 1 Then _
            g_show.TrafficStart "FindHome"
    End If
    
    check_connected
    HW_Move (g_dSetPark)
    'Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
        "Method FindHome" & MSG_NOT_IMPLEMENTED
                    
End Sub

Public Sub OpenShutter()

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "OpenShutter"
    End If

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
        "Method OpenShutter" & MSG_NOT_IMPLEMENTED
    
End Sub

Public Sub Park()
 
    If Not g_show Is Nothing Then
        If g_show.chkSlew.Value = 1 Then _
            g_show.TrafficStart "Park"
    End If
    
    check_connected
    
    If Not g_bAtPark Then _
        HW_Park
        
End Sub

Public Sub SetPark()

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "SetPark: " & FmtSexa(g_dDomeAz, False)
    End If
    
    check_connected
    g_dSetPark = g_dDomeAz
    g_bAtPark = True
    g_handBox.RefreshLEDs
    g_handBox.LabelButtons

    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
    
End Sub

Public Sub SetupDialog()
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficStart "SetupDialog"
    End If
    
    DoSetup
    
    If Not g_show Is Nothing Then
        If g_show.chkOther.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
    
End Sub

Public Sub SlewToAltitude(ByVal Alt As Double)

    If Not g_show Is Nothing Then
        If g_show.chkSlew.Value = 1 Then _
            g_show.TrafficLine "SlewToAltitude:" & FmtSexa(Alt, False)
    End If
    
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
        "Method SlewToAltitude" & MSG_NOT_IMPLEMENTED
        
End Sub

Public Sub SlewToAzimuth(ByVal Az As Double)

    If Not g_show Is Nothing Then
        If g_show.chkSlew.Value = 1 Then _
            g_show.TrafficLine "SlewToAzimuth: " & FmtSexa(Az, False)
    End If
    
    check_connected
                
    check_Az Az
    HW_Move Az
    
    ' slew code in Startup.bas journals end of slew
End Sub

Public Sub SyncToAzimuth(ByVal Az As Double)
      
    If Not g_show Is Nothing Then
        If g_show.chkSlew.Value = 1 Then _
            g_show.TrafficLine "SyncToAzimuth: " & FmtSexa(Az, False)
    End If
    
    check_connected
    check_Az Az
    HW_Sync Az
    
    If Not g_show Is Nothing Then
        If g_show.chkSlew.Value = 1 Then _
            g_show.TrafficEnd " (done)"
    End If
 
End Sub

' ----------------
' Private Routines
' ----------------

'---------------------------------------------------------------------
'
' check_connected() - Raise an error if the dome is not connected
'
'---------------------------------------------------------------------
Private Sub check_connected()

    If Not g_bConnected Then _
        Err.Raise SCODE_NOT_CONNECTED, ERR_SOURCE, MSG_NOT_CONNECTED
        
End Sub

'---------------------------------------------------------------------
'
' check_Az() - Raise an error if Az is bogus
'
'---------------------------------------------------------------------
Private Sub check_Az(Az As Double)

    If Az = INVALID_COORDINATE Then _
        Err.Raise SCODE_NO_TARGET_COORDS, ERR_SOURCE, _
            "Azimuth " & MSG_NO_TARGET_COORDS
     
    If Az > 360 Or Az < -360 Then _
        Err.Raise SCODE_VAL_OUTOFRANGE, ERR_SOURCE, _
            "Azimuth " & MSG_VAL_OUTOFRANGE
            
End Sub


