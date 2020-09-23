VERSION 5.00
Begin VB.Form frmAcercaDe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ODBC Tools"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcercaDe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   4035
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Septiembre 2002"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Correo electrónico: arturoruzc@hotmail.com"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1860
         Width           =   3090
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sitio web: www.programadores.ma.cx"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1620
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Creado por: José Arturo Ruz Carrillo."
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1380
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "La utilización de esta librería es totalmente gratuita."
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1020
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   $"frmAcercaDe.frx":08CA
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3840
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Versión 1.0"
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ODBC Tool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00925F56&
      Height          =   675
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   -60
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ODBC Tool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   675
      Index           =   1
      Left            =   150
      TabIndex        =   8
      Top             =   -30
      Width           =   2655
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
