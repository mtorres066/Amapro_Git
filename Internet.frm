VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Internet 
   BackColor       =   &H80000001&
   Caption         =   "Internet Erick"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Internet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Internet.frx":628A
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Barra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   8340
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "DESP"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "12:28"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "17/11/2010"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7594
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdBotones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1800
      Picture         =   "Internet.frx":6B54
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Actualizar"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton CmdBotones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1200
      Picture         =   "Internet.frx":6F96
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Detener"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton CmdBotones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   0
      Picture         =   "Internet.frx":73D8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Atras"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton CmdBotones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   600
      Picture         =   "Internet.frx":781A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Adelante"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Text            =   "www.fepsa.com.mx"
      ToolTipText     =   "escriba direccion de internet"
      Top             =   360
      Width           =   8295
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      DragIcon        =   "Internet.frx":7C5C
      Height          =   9015
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "ventana de internet"
      Top             =   720
      Width           =   11895
      ExtentX         =   20981
      ExtentY         =   15901
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   870
   End
End
Attribute VB_Name = "Internet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next

        If Index = 0 Then
            WebBrowser1.GoBack
        ElseIf Index = 1 Then
            WebBrowser1.GoForward
        ElseIf Index = 2 Then
            WebBrowser1.Stop
        ElseIf Index = 3 Then
            WebBrowser1.Refresh
        End If
        
        
        If Err <> 0 Then
            MsgBox Err.Number & " " & Err.Description
        End If
End Sub

Private Sub Form_Load()
On Error Resume Next
            MousePointer = 11
                WebBrowser1.Navigate Text1.Text
            MousePointer = 0
            If Err <> 0 Then
                    MsgBox Err.Number & " " & Err.Description
            End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
            WebBrowser1.Height = Me.ScaleHeight - 1100
            WebBrowser1.Width = Me.ScaleWidth - 100
            
            If Err <> 0 Then
            End If
End Sub

Private Sub Text1_GotFocus()
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
            If KeyAscii = 13 Then
                MousePointer = 11
                    WebBrowser1.Navigate Text1.Text
                MousePointer = 0
            End If
            
            If Err <> 0 Then
                    MsgBox Err.Number & " " & Err.Description
            End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
        MousePointer = 11
                Barra.Panels.Item(6).Text = "descargando pagina"
        MousePointer = 0
        
End Sub

Private Sub WebBrowser1_DownloadComplete()
        MousePointer = 11
                Barra.Panels.Item(6).Text = "pagina completa"
        MousePointer = 0
End Sub

