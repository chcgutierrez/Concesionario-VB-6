VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_bPais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Pais"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnBsqAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dtgPaisAct 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cod_pais"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nom_pais"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "est_pais"
         Caption         =   "Estado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "obs_gen"
         Caption         =   "Observaciones"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Realizar B�squeda"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6880
      Begin VB.CommandButton btnBusqPais 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   5730
         TabIndex        =   3
         Top             =   280
         Width           =   975
      End
      Begin VB.TextBox txtDescPais 
         Height          =   375
         Left            =   1250
         TabIndex        =   2
         Top             =   300
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   840
      End
   End
End
Attribute VB_Name = "frm_bPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bsqRespuesta As Boolean

Public Function BusquedaPais(ByRef strCodPais As String, ByRef strDescPais As String) As Boolean
       
    On Error GoTo ControlError
    
        Set dtgPaisAct.DataSource = Nothing
        bsqRespuesta = False
        Me.Show vbModal
        If bsqRespuesta Then
            If Not dtgPaisAct.DataSource Is Nothing Then 'Grid del form busqueda
                strCodPais = dtgPaisAct.Columns(0).Text
                strDescPais = dtgPaisAct.Columns(1).Text
                BusquedaPais = True
            Else
                BusquedaPais = False
            End If
        Else
            BusquedaPais = False
        End If
        
    Exit Function
ControlError:
MsgBox "Ha ocurrido un error en la aplicaci�n." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripci�n del error: " & Err.Description, vbCritical, App.Title
          
End Function

Private Sub btnBsqAceptar_Click()
On Error GoTo ControlError
    bsqRespuesta = True
    Me.Hide
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicaci�n." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripci�n del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub btnBusqPais_Click()

Dim rsBusqPais As ADODB.Recordset
    
    On Error GoTo ControlError
    
    If Len(Me.txtDescPais.Text) > 0 Then
        Set rsBusqPais = TraerPaisDesc(txtDescPais.Text)
            If rsBusqPais.RecordCount > 0 Then
                Set dtgPaisAct.DataSource = rsBusqPais
                dtgPaisAct.Columns("Codigo").Width = 900
                dtgPaisAct.Columns("Nombre").Width = 2300
                dtgPaisAct.Columns("Estado").Width = 800
                dtgPaisAct.Columns("Observaciones").Width = 1300
            Else
                Set dtgPaisAct.DataSource = Nothing
                Me.txtDescPais.SelStart = 0
                Me.txtDescPais.SelLength = Len(Me.txtDescPais.Text)
                MsgBox "El pais no existe o est� inactivo", vbOKOnly, "Buscar Pais"
'                Cancel
                Exit Sub
            End If
    Else
         MsgBox "Debe ingresar un criterio para realizar la busqueda.", vbOKOnly, "Criterio Inv�lido"
         Me.txtDescPais.SetFocus
         Set dtgPaisAct.DataSource = Nothing
'         Cancel = True
         Exit Sub
    End If
    
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicaci�n." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripci�n del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub dtgPaisAct_DblClick()
On Error GoTo ControlError
    bsqRespuesta = True
    Me.Hide
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicaci�n." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripci�n del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub txtDescPais_Change()
On Error GoTo ControlError
    txtDescPais.Text = UCase(txtDescPais.Text)
    txtDescPais.SelStart = Len(txtDescPais)
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicaci�n." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripci�n del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub
