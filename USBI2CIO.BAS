Attribute VB_Name = "modUsbI2cIoAPI"
Option Explicit
' This module file contains the necessary definitions, types, and declarations for accessing the
' functions provided by the UsbI2cIo API dll.

' Global constants and enumerations definitions
Public Const INVALID_HANDLE_VALUE As Integer = -1
' UsbI2cIo maximum devices
Public Const USBI2CIO_MAX_DEVICES As Byte = 127
'Public Const USBI2CIO_DLL_NAME As String = "UsbI2cIo.dll"

' I2C transaction constants
Public Const I2C_HEADER_SIZE As Byte = 6
Public Const I2C_MAX_DATA As Byte = 255

' I2C Transaction Types
Public Enum I2C_TRANS_TYPE
   I2C_TRANS_NOADR = 0                      ' for I2C devices that do not use sub-addresses
   I2C_TRANS_8ADR = 1                       ' for I2C devices that use 8 bit sub-addresses
   I2C_TRANS_16ADR = 2                      ' for I2C devices that use 16 bit sub-addresses
End Enum


' Global type definitions for UsbI2cIo API DLL (correspond to values in UsbI2cIo.h)

Public Type Word          ' provides easy access to high and low bytes of two-byte entity
  lo As Byte
  hi As Byte
End Type

Public Type I2C_TRANS     ' I2C Transaction Structure, used to specify I2C transaction info
  byType As Byte          ' see I2C_TRAN_TYPE enum (above)
  byDevId As Byte         ' bits 7-1 = the I2C device ID, bit 0 is auto set/cleared by call
  wMemAddr As Word        ' if accessing a device with sub-addressing, sub-address goes here
  wCount As Word          ' count of bytes in Data array
  Data(64) As Byte        ' I2C transaction data
End Type

Public Type DEVINFO
  byInstance As Byte      ' instance number of device
  SerialId(8) As Byte     ' 8 bytes Serial ID string of device and a NULL termination
End Type                  ' Note: in Vb, the array size is 1 greater than number specified


' UsbI2cIo API DLL function declarations

Declare Function DAPI_GetDllVersion Lib "UsbI2cIo.dll" () As Word

'Declare Function DAPI_GetDriverVersion Lib "UsbI2cIo.dll" () As Word

'Declare Function DAPI_GetFirmwareVersion Lib "UsbI2cIo.dll" () As Word

Declare Function DAPI_GetDeviceCount Lib "UsbI2cIo.dll" ( _
  ByVal lpsDevName As String) _
  As Byte

Declare Function DAPI_GetDeviceInfo Lib "UsbI2cIo.dll" ( _
  ByVal lpsDevName As String, _
  ByRef lpDevInfo As DEVINFO) _
  As Byte

Declare Function DAPI_GetSerialId Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByVal lpsSerialId As String) _
  As Byte
  
Declare Function DAPI_DetectDevice Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long) _
  As Boolean

Declare Function DAPI_OpenDeviceInstance Lib "UsbI2cIo.dll" ( _
  ByVal lpsDevName As String, _
  ByVal byDevInstance As Byte) _
  As Long

Declare Function DAPI_CloseDeviceInstance Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long) _
  As Boolean

Declare Function DAPI_OpenDeviceBySerialId Lib "UsbI2cIo.dll" ( _
  ByVal lpsDevName As String, _
  ByVal lpsSerialId As String) _
  As Long

Declare Function DAPI_GetIoConfig Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByRef pulIoPortData As Long) _
  As Boolean

Declare Function DAPI_ConfigIoPorts Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByVal ulIoPortConfig As Long) _
  As Boolean

Declare Function DAPI_ReadIoPorts Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByRef pulIoPortData As Long) _
  As Boolean
  
Declare Function DAPI_WriteIoPorts Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByVal ulIoPortData As Long, _
  ByVal ulIoPortMask As Long) _
  As Boolean

Declare Function DAPI_ReadI2c Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByRef TransI2c As I2C_TRANS) _
  As Long

Declare Function DAPI_WriteI2c Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByRef TransI2c As I2C_TRANS) _
  As Long

Declare Function DAPI_ReadDebugBuffer Lib "UsbI2cIo.dll" ( _
  ByRef DebugBuf As Byte, _
  ByVal hDevInstance As Long, _
  ByVal ulMaxBytes As Long) _
  As Long

