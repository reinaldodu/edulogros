VERSION 5.00
Begin VB.Form CONS_PRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de profesor"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   Icon            =   "CONS_PRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   320
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   120
      Picture         =   "CONS_PRO.frx":0442
      ScaleHeight     =   3795
      ScaleWidth      =   1635
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&IMPRIMIR"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "CONSULTA DE PROFESOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3615
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   7455
      Begin VB.TextBox Text11 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1305
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "FECHA DE NACIMIENTO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ESCALAFON:"
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
         Left            =   3600
         TabIndex        =   23
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TITULO :"
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
         Left            =   3600
         TabIndex        =   20
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label7 
         Caption         =   "AÑO DE INGRESO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TELEFONO:"
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
         Left            =   3600
         TabIndex        =   18
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DIRECCION:"
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
         Left            =   3600
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FACTOR R-H:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CEDULA No:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "APELLIDOS:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRES:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   990
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "PROFESOR No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7320
      TabIndex        =   22
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "CONS_PRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim ini As inicio
RESP = MsgBox("DESEA IMPRIMIR LA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
    Close #NAR
    Printer.ScaleMode = 7
    Printer.CurrentY = 1
    Printer.CurrentX = 8
    Printer.Font.Size = 12
    Printer.Print "DATOS DEL PROFESOR"
    Printer.Font.Size = 10
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print ini.nombre
    Printer.CurrentY = 4
    Printer.CurrentX = 2
    Printer.Print "PROFESOR No." & Text9.Text;
    Printer.CurrentY = 5
    Printer.CurrentX = 2
    Printer.Print "NOMBRES: ";
    Printer.Print Text1.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "APELLIDOS: ";
    Printer.Print Text2.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "DOCUMENTO: ";
    Printer.Print Text3.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "FACTOR R-H: ";
    Printer.Print Text4.Text
    Printer.CurrentY = 5
    Printer.CurrentX = 9.5
    Printer.Print "DIRECCION: ";
    Printer.Print Text5.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 9.5
    Printer.Print "TELEFONO: ";
    Printer.Print Text6.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 9.5
    Printer.Print "AÑO DE INGRESO: ";
    Printer.Print Text7.Text;
    Printer.CurrentX = 15
    Printer.Print "ESCALAFON: ";
    Printer.Print Text10.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 9.5
    Printer.Print "ESPECIALIZACION: ";
    Printer.Print Text8.Text
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra toda la información de un profesor."
End Sub
