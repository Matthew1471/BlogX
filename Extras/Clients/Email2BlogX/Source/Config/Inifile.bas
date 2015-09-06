Attribute VB_Name = "Inifile"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

' Read INI
Public Function ReadIni(Filename As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
ReadIni = Left(RetVal, v)
End Function
