VERSION 5.00
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PharmaStar"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13680
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   48
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   2190
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "frm_main.frx":628A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16140
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

