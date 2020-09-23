VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MySQL"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private nRet As Long
Private sAtributos, m_DATABASE, m_DESC As String
Private m_DATASOURCE, m_UID, m_PASSWORD, m_SERVER, m_PORT As String

Private cnnConexion As ADODB.Connection

Private Sub Class_Initialize()
    nRet = 0&
    sAtributos = ""
    m_DATASOURCE = ""
    m_DESC = "Creado por ODBC Tool"
    m_DATABASE = ""
    m_UID = ""
    m_PASSWORD = ""
    m_SERVER = ""
    m_PORT = "3306"
End Sub

' ==================================================
' P R O P I E D A D E S
' ==================================================

Public Property Let Database(ByVal sDatabase As String)
    m_DATABASE = sDatabase
End Property
Public Property Get Database() As String
    Database = m_DATABASE
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

Public Property Let Server(ByVal sServer As String)
    m_SERVER = sServer
End Property
Public Property Get Server() As String
    Server = m_SERVER
End Property

Public Property Let Port(ByVal sPort As String)
    m_PORT = sPort
End Property
Public Property Get Port() As String
    Port = m_PORT
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
    If m_SERVER = "" Then: Crear = "No se ha definido la direccion del servidor": Exit Function
    If m_PORT = "" Then: Crear = "No se ha definido el puerto del servidor": Exit Function
    If m_DATABASE = "" Then: Crear = "No se ha definido la base de datos": Exit Function
    If m_UID = "" Then: Crear = "No se ha definido el nombre de usuario": Exit Function
    
    sAtributos = "SERVER=" & m_SERVER & Chr$(0)
    sAtributos = sAtributos & "PORT=" & m_PORT & Chr$(0)
    sAtributos = sAtributos & "DATABASE=" & m_DATABASE & Chr$(0)
    sAtributos = sAtributos & "UID=" & m_UID & Chr$(0)
    sAtributos = sAtributos & "PASSWORD=" & m_PASSWORD & Chr$(0)
    sAtributos = sAtributos & "DESC=" & m_DESC & Chr$(0)
    sAtributos = sAtributos & "DSN=" & m_DATASOURCE & Chr$(0)
    nRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYSTEM_DSN, ODBC_DRIVER_MYSQL, sAtributos)
End Function

' =======================================
' Eliminar(): Elimina un DSN del ODBC
' =======================================
Public Function Eliminar() As String
    Eliminar = ""
    If m_DATASOURCE = "" Then: Eliminar = "No se ha definido el nombre del origen ODBC": Exit Function
    sAtributos = sAtributos & "DSN=" & m_DATASOURCE & Chr$(0)
    nRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_SYSTEM_DSN, ODBC_DRIVER_MYSQL, sAtributos)
End Function

' =========================================
' Actualizar(): Modifica un DNS del ODBC
' =========================================
Public Function Actualizar() As String
    Actualizar = ""
    If m_DATASOURCE = "" Then: Actualizar = "No se ha definido el nombre del origen ODBC": Exit Function
    If m_SERVER = "" Then: Actualizar = "No se ha definido la direccion del servidor": Exit Function
    If m_PORT = "" Then: Actualizar = "No se ha definido el puerto del servidor": Exit Function
    If m_DATABASE = "" Then: Actualizar = "No se ha definido la base de datos": Exit Function
    If m_UID = "" Then: Actualizar = "No se ha definido el nombre de usuario": Exit Function
    
    sAtributos = "SERVER=" & m_SERVER & Chr$(0)
    sAtributos = sAtributos & "PORT=" & m_PORT & Chr$(0)
    sAtributos = sAtributos & "DATABASE=" & m_DATABASE & Chr$(0)
    sAtributos = sAtributos & "UID=" & m_UID & Chr$(0)
    sAtributos = sAtributos & "PASSWORD=" & m_PASSWORD & Chr$(0)
    sAtributos = sAtributos & "DESC=" & m_DESC & Chr$(0)
    sAtributos = sAtributos & "DSN=" & m_DATASOURCE & Chr$(0)
    nRet = SQLConfigDataSource(vbAPINull, ODBC_CONFIG_SYSTEM_DSN, ODBC_DRIVER_MYSQL, sAtributos)
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
