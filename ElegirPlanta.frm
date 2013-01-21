VERSION 5.00
Begin VB.Form ElegirPlanta 
   BackColor       =   &H80000003&
   Caption         =   "Elegir Planta"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptOpc 
      BackColor       =   &H80000003&
      Caption         =   "Chiapas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   975
      Left            =   1440
      Picture         =   "ElegirPlanta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.OptionButton OptOpc 
      BackColor       =   &H80000003&
      Caption         =   "San Luis Potosi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.OptionButton OptOpc 
      BackColor       =   &H80000003&
      Caption         =   "Fepsa Culiacan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      Caption         =   "De Que Planta Importara Los Datos ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3315
   End
End
Attribute VB_Name = "ElegirPlanta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    If OptOpc.Item(0).Value = True Then
        GPlanta = "CULIACAN"
    ElseIf OptOpc.Item(1).Value = True Then
        GPlanta = "SAN LUIS"
    ElseIf OptOpc.Item(2).Value = True Then
        GPlanta = "CHIAPAS"
    End If
        Unload Me
        GeneraCapturaRutinas.Show
        
        
    If Err <> 0 Then
    End If
End Sub
