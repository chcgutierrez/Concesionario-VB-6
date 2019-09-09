Attribute VB_Name = "mod_Declaraciones"
Option Explicit

Global ConexSQL As New ADODB.Connection 'variable global para la conexion a la BD
Global RS_tblUSUARIO As New ADODB.Recordset 'variable global para la conexion a la Tabla
Global cmdSQL As New ADODB.Command
Global rstSQL As New ADODB.Recordset
