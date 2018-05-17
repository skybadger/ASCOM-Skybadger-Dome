Attribute VB_Name = "modWinAPI"
Option Explicit
'
' This module file contains definitions, types, and declarations for accessing the
' functions provided by the Windows Win32 API.
'


' Win32 API constant declarations
Public Const INVALID_HANDLE_VALUE = -1



' Win32 API function declarations

Declare Function GetLastError Lib "kernel32" () As Long

Declare Function DeviceIoControl Lib "kernel32" ( _
  ByVal hDevice As Long, _
  ByVal dwIoControlCode As Long, _
  lpInBuffer As Byte, _
  ByVal nInBufferSize As Long, _
  lpOutBuffer As Byte, _
  ByVal nOutBufferSize As Long, _
  lpBytesReturned As Long, _
  lpOverlapped As Long) _
  As Long

Declare Function WriteFile Lib "kernel32" ( _
  ByVal hFile As Long, _
  lpBuffer As Any, _
  ByVal nNumberOfBytesToWrite As Long, _
  lpNumberOfBytesWritten As Long, _
  lpOverlapped As Long) _
  As Long

Declare Function ReadFile Lib "kernel32" ( _
  ByVal hFile As Long, _
  lpBuffer As Any, _
  ByVal nNumberOfBytesToRead As Long, _
  lpNumberOfBytesRead As Long, _
  lpOverlapped As Long) _
  As Long


' Error Code Defines from API

'   Code    Description                                             Name
 
'   0       The operation completed successfully.                   ERROR_SUCCESS
'   1       Incorrect function.                                     ERROR_INVALID_FUNCTION
'   2       The system cannot find the file specified.              ERROR_FILE_NOT_FOUND
'   3       The system cannot find the path specified.              ERROR_PATH_NOT_FOUND
'   4       The system cannot open the file.                        ERROR_TOO_MANY_OPEN_FILES
'   5       Access is denied.                                       ERROR_ACCESS_DENIED
'   6       The handle is invalid.                                  ERROR_INVALID_HANDLE
'
