VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Iniciar Sesión"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "frm_login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_cancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      MousePointer    =   4  'Icon
      TabIndex        =   6
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton btn_login 
      Caption         =   "Iniciar Sesión"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      MousePointer    =   4  'Icon
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox txt_clave 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox txt_usuario 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar Sesión"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   5
      Top             =   1320
      Width           =   4350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2880
      TabIndex        =   1
      Top             =   3720
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   3480
      TabIndex        =   0
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Picture         =   "frm_login.frx":058A
      Top             =   0
      Width           =   12780
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_cancelar_Click()
'cancelar
End Sub

Private Sub btn_login_Click()
'login
'buscar
End Sub

Private Sub Form_Load()
'Abrir_tblUSUARIO
End Sub

Private Sub txt_clave_GotFocus()
txt_clave.BackColor = &H8000000A
End Sub

Private Sub txt_clave_LostFocus()
txt_clave.BackColor = &H80000005
End Sub

Private Sub txt_clave_Validate(Cancel As Boolean)

'Variable Global en Mod_Ppal
'g_strConexion = "Provider=SQLOLEDB.1;" & _
'           "Persist Security Info=False;" & _
'           "User ID=ccgutierrezm;" & _
'           "PWD=1030538949;" & _
'           "Initial Catalog=almCarros;" & _
'           "Data Source=CLIENTE-PC;" & _
'           "PP=fc:YYYY/MM/DD@fl:YYYY/MM/DD HH:NN:SS AM/PM@+:+@isnull:IsNull@;" & _
'           "Database=almCarros;" & _
'           "AnsiNPW=no"
End Sub

Private Sub txt_usuario_GotFocus()
txt_usuario.BackColor = &H8000000A
End Sub

Private Sub txt_usuario_LostFocus()
txt_usuario.BackColor = &H80000005
End Sub
