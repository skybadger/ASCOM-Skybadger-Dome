Attribute VB_Name = "Rotate"
'VB-Helper DLL, Version 2.02
'Copyright (c) 1996-97 SoftCircuits Programming(R)
'Redistributed by Permission.
'
'This package includes a helper DLL for 32-bit Visual Basic. This DLL
'provides a number of routines that perform tasks that are either
'difficult or impossible to accomplish in Visual Basic alone. Some
'sample programs are also provided to demonstrate use of the DLL.
'Please see the included help file for details on all of the routines
'included within the DLL.
'
'The VB-Helper DLL is freeware that you can use freely with your own
'programs. Any portion of the sample programs may also be incorporated
'into your own applications. However, you may only distribute
'Vbhlp32.dll as a) part of your own application that uses this DLL or
'b) within this complete and unmodified package (i.e., you may
'distribute the entire Vbhlp32.zip file).
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623
Option Explicit

Declare Function vbGetAddress Lib "VBHLP32.DLL" (pData As Any) As Long
Declare Sub vbFillMemory Lib "VBHLP32.DLL" (pDest As Any, ByVal nValue As Byte, ByVal nCount As Long)
Declare Sub vbCopyMemory Lib "VBHLP32.DLL" (pDest As Any, pSource As Any, ByVal nCount As Long)
Declare Function vbLoByte Lib "VBHLP32.DLL" (ByVal nValue As Integer) As Integer
Declare Function vbHiByte Lib "VBHLP32.DLL" (ByVal nValue As Integer) As Integer
Declare Function vbLoWord Lib "VBHLP32.DLL" (ByVal nValue As Long) As Integer
Declare Function vbHiWord Lib "VBHLP32.DLL" (ByVal nValue As Long) As Integer
Declare Function vbMakeWord Lib "VBHLP32.DLL" (ByVal nLoByte As Integer, ByVal nHiByte As Integer) As Integer
Declare Function vbMakeLong Lib "VBHLP32.DLL" (ByVal nLoWord As Integer, ByVal nHiWord As Integer) As Long
Declare Function vbShiftRight Lib "VBHLP32.DLL" (ByVal nValue As Integer, ByVal nBits As Integer) As Integer
Declare Function vbShiftLeft Lib "VBHLP32.DLL" (ByVal nValue As Integer, ByVal nBits As Integer) As Integer
Declare Function vbShiftRightLong Lib "VBHLP32.DLL" (ByVal nValue As Long, ByVal nBits As Integer) As Long
Declare Function vbShiftLeftLong Lib "VBHLP32.DLL" (ByVal nValue As Long, ByVal nBits As Integer) As Long
Declare Function vbRotateRight Lib "VBHLP32.DLL" (ByVal nValue As Integer, ByVal nBits As Integer) As Integer
Declare Function vbRotateLeft Lib "VBHLP32.DLL" (ByVal nValue As Integer, ByVal nBits As Integer) As Integer
Declare Function vbRotateRightLong Lib "VBHLP32.DLL" (ByVal nValue As Long, ByVal nBits As Integer) As Long
Declare Function vbRotateLeftLong Lib "VBHLP32.DLL" (ByVal nValue As Long, ByVal nBits As Integer) As Long
Declare Sub vbPackUDT Lib "VBHLP32.DLL" (pUDT As Any, ppResult As Long, ByVal pszFields As String)
Declare Sub vbUnpackUDT Lib "VBHLP32.DLL" (pUDT As Any, ppResult As Long)
Declare Function vbPackUDTGetSize Lib "VBHLP32.DLL" (ppResult As Long) As Long
Declare Sub vbPackUDTFree Lib "VBHLP32.DLL" (ppResult As Long)
Declare Function vbGetHelperVersion Lib "VBHLP32.DLL" () As Integer
Declare Sub vbShowAboutBox Lib "VBHLP32.DLL" ()

