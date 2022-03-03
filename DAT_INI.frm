VERSION 5.00
Begin VB.Form DAT_INI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos iniciales"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5085
   Icon            =   "DAT_INI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Variables"
      Height          =   465
      Left            =   3720
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   465
      Left            =   2160
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   4815
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1320
         TabIndex        =   15
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1320
         TabIndex        =   4
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SECRETARIO:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "RECTOR:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "AÑO ACTUAL:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "INSTITUCION:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "OPCIONAL :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RESOLUCION :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1140
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DATOS INICIALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1800
      Picture         =   "DAT_INI.frx":0442
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "DAT_INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim ini As inicio
If RTrim(Text1.Text) = "" Then
    MsgBox "ESCRIBA LA RESOLUCIÓN DE APROBACIÓN", 16, "DATOS INICIALES"
    Text1.SetFocus
    Exit Sub
End If
If RTrim(Text2.Text) = "" Then
    MsgBox "ESCRIBA EL NOMBRE DE LA INSTITUCIÓN", 16, "DATOS INICIALES"
    Text2.SetFocus
    Exit Sub
End If
If RTrim(Text4.Text) = "" Then
    MsgBox "ESCRIBA EL AÑO ACADÉMICO ACTUAL", 16, "DATOS INICIALES"
    Text4.SetFocus
    Exit Sub
End If
If RTrim(Text5.Text) = "" Then
    MsgBox "ESCRIBA EL NOMBRE DEL RECTOR", 16, "DATOS INICIALES"
    Text5.SetFocus
    Exit Sub
End If
If RTrim(Text6.Text) = "" Then
    MsgBox "ESCRIBA EL NOMBRE DEL SECRETARIO", 16, "DATOS INICIALES"
    Text6.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "inicial.edu" For Output As #NAR
ini.ciudad = Text1.Text
ini.nombre = Text2.Text
ini.modalidad = Text3.Text
ini.Telefono = Text4.Text
ini.Rector = Text5.Text
ini.secretario = Text6.Text
Write #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
ENTRADA.Caption = "EDULOGROS - " & ini.nombre
Unload Me
End Sub

Private Sub Command2_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    VarEdulogros.Show 1
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Los Datos Iniciales sirven para identificar a la institución en los reportes impresos por el programa."
End Sub

Private Sub Form_Load()
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
'Text1.MaxLength = 20
'Text2.MaxLength = 33
'Text3.MaxLength = 15
'Text4.MaxLength = 30
'Text5.MaxLength = 80
Text1 = Trim(ini.ciudad)
Text2 = Trim(ini.nombre)
Text3 = Trim(ini.modalidad)
Text4 = Trim(ini.Telefono)
Text5 = Trim(ini.Rector)
Text6 = Trim(ini.secretario)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
