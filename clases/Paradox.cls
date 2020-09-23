VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Paradox"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private nRet As Long
Private sAtributos, m_DBQ, m_DESC As String
Private m_DATASOURCE, m_UID, m_PASSWORD, m_DEFAULTDIR, m_FIL As String

Private cnnConexion As ADODB.Connection

Private Sub Class_Initialize()
    nRet = 0&
    sAtributos = ""
    m_DATASOURCE = ""
    m_DESC = "Creado por ODBC Tool"
    m_UID = "admin"
    m_PASSWORD = ""
    m_FIL = "Paradox 5.X"
    m_DEFAULTDIR = ""
    m_DBQ = ""
End Sub

' ==================================================
' P R O P I E D A D E S
' ==================================================

Public Property Let Database(ByVal sDatabase As String)
    m_DBQ = sDatabase
End Property
Public Property Get DBQ() As String
    DBQ = m_DBQ
End Property

Public Property Let DataSourceName(ByVal sDataSourceName As String)
    m_DATASOURCE = sDataSourceName
End Property
Public Property Get DataSourceName() As String
    DataSourceName = m_DATASOURCE
End Property

Public Property Let UID(ByVal sUID As String)
    m_UID = sUID
End Property
Public Property Get UID() As String
    UID = m_UID
End Property

Public Property Let Password(ByVal sPassword As String)
    m_PASSWORD = sPassword
End Property
Public Property Get Password() As String
    Password = m_PASSWORD
End Property

Public Property Let DefaultDir(ByVal sDefaultDir As String)
    m_DEFAULTDIR = sDefaultDir
End Property
Public Property Get DefaultDir() As String
    DefaultDir = m_DEFAULTDIR
End Property

Public Property Let Desc(ByVal sDesc As String)
    m_DESC = sDesc
End Property
Public Property Get Desc() As String
    Desc = m_DESC
End Property


' ==================================================
' M E T O D O S
' ==================================================

' ===============================
' Crear() Crea un DSN en ODBC
' ===============================
Public Function Crear() As String
    Crear = ""
    If m_DATASOURCE = "" Then: Crear = "No se ha definido el nombre del origen ODBC": Exit Function
    If m_DBQ = "" Then: Crear = "No se ha definido la base de datos": Exit Function
    
    sAtributos = "DEFAULTDIR=" & m_DEFAULTDIR & Chr$(0)
    sAtributos = sAtributos & "DBQ=" & m_DBQ & Chr$(0)
    sAtributos = sAtributos & "UID=" & m_UID & Chr$(0)
    sAtributos = sAtributos & "PASSWORD=" & m_PASSWORD & Chr$(0)
    sAtributos = sAtributos & "DESC=" & m_DESC & Chr$(0)
    sAtributos = sAtributos & "DSN=" & m_DATASOURCE & Chr$(0)
    sAtributos = sAtributos & "FIL=" & m_FIL & Chr$(0)
    nRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYSTEM_DSN, ODBC_DRIVER_PARADOX, sAtributos)
End Function

' =======================================
' Eliminar(): Elimina un DSN del ODBC
' =======================================
Public Function Eliminar() As String
    Eliminar = ""
    If m_DATASOURCE = "" Then: Eliminar = "No se ha definido el nombre del origen ODBC": Exit Function
    sAtributos = sAtributos & "DSN=" & m_DATASOURCE & Chr$(0)
    nRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_SYSTEM_DSN, ODBC_DRIVER_PARADOX, sAtributos)
End Function

' =========================================
' Actualizar(): Modifica un DNS del ODBC
' =========================================
Public Function Actualizar() As String
    Actualizar = ""
    If m_DATASOURCE = "" Then: Actualizar = "No se ha definido el nombre del origen ODBC": Exit Function
    If m_DBQ = "" Then: Actualizar = "No se ha definido la base de datos": Exit Function
    
    sAtributos = "DEFAULTDIR=" & m_DEFAULTDIR & Chr$(0)
    sAtributos = sAtributos & "DBQ=" & m_DBQ & Chr$(0)
    sAtributos = sAtributos & "UID=" & m_UID & Chr$(0)
    sAtributos = sAtributos & "PASSWORD=" & m_PASSWORD & Chr$(0)
    sAtributos = sAtributos & "DESC=" & m_DESC & Chr$(0)
    sAtributos = sAtributos & "DSN=" & m_DATASOURCE & Chr$(0)
    sAtributos = sAtributos & "FIL=" & m_FIL & Chr$(0)
    nRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYSTEM_DSN, ODBC_DRIVER_PARADOX, sAtributos)
End Function

' ===============================
' Conectar(): Crea una conexion
' ===============================
Public Function Conectar() As ADODB.Connection
    Dim NuevaConexion As ADODB.Connection
    Set NuevaConexion = New ADODB.Connection
    NuevaConexion.ConnectionString = "Data Source=" & m_DATASOURCE
    NuevaConexion.Open
    Set Conectar = NuevaConexion
End Function

' ===================================
' Desconectar(): Cierra una conexion
' ===================================
Public Sub Desconectar(Conexion As ADODB.Connection)
    Conexion.Close
End Sub

' ==================================================
' SQLQuery(): Almacena una consulta en un RecordSet
' ==================================================
Public Function SQLQuery(Conexion As ADODB.Connection, cQueryString As String) As ADODB.Recordset
    Set SQLQuery = Conexion.Execute(cQueryString)
End Function

' =====================================
' SQLExec(): Ejecuta una sentencia SQL
' =====================================
Public Sub SQLExec(Conexion As ADODB.Connection, cQueryString As String)
    Conexion.Execute cQueryString
End Sub

' ==============================================
' AcercaDe() Muestra el diálogo de Acerca de...
' ==============================================
Public Sub AcercaDe()
    frmAcercaDe.Show
End Sub
