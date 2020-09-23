Attribute VB_Name = "Declaraciones"
Option Explicit

Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
    (ByVal hwndParent As Long, ByVal fRequest As Long, _
    ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long


Global Const ODBC_ADD_DSN = 1
Global Const ODBC_CONFIG_DSN = 2
Global Const ODBC_REMOVE_DSN = 3
Global Const ODBC_ADD_SYSTEM_DSN = 4
Global Const ODBC_CONFIG_SYSTEM_DSN = 5
Global Const ODBC_REMOVE_SYSTEM_DSN = 6
Global Const vbAPINull As Long = 0&

Global Const ODBC_DRIVER_ACCESS = "Microsoft Access Driver (*.mdb)"
Global Const ODBC_DRIVER_DBASE = "Microsoft dBase Driver (*.dbf)"
Global Const ODBC_DRIVER_PARADOX = "Microsoft Paradox Driver (*.db )"
Global Const ODBC_DRIVER_MYSQL = "MySQL ODBC 3.51 Driver"
