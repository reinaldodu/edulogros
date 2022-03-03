VERSION 5.00
Begin VB.Form CONS_ALUM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de alumno"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   Icon            =   "CONS_ALUM.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&IMPRIMIR"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "------------------------------------------CONSULTA  DE ESTUDIANTE-----------------------------------"
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
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox Text23 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL:"
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
         TabIndex        =   48
         Top             =   4080
         Width           =   630
      End
      Begin VB.Image picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1785
         Left            =   7560
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "TEL.CASA:"
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
         TabIndex        =   47
         Top             =   3000
         Width           =   960
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "E.P.S:"
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
         TabIndex        =   46
         Top             =   3720
         Width           =   555
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
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
         Left            =   4800
         TabIndex        =   45
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "TEL:"
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
         Left            =   3120
         TabIndex        =   44
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "MADRE:"
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
         TabIndex        =   43
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "TEL:"
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
         Left            =   3120
         TabIndex        =   42
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "PADRE:"
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
         TabIndex        =   41
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "EDAD:"
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
         Left            =   7920
         TabIndex        =   40
         Top             =   2040
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "(dd/mm/aaaa)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6120
         TabIndex        =   39
         Top             =   480
         Width           =   1020
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "CONS_ALUM.frx":0442
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "SEXO:"
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
         Left            =   7920
         TabIndex        =   38
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "MATRICULA:"
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
         Left            =   2760
         TabIndex        =   37
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "GRADO:"
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
         Left            =   4800
         TabIndex        =   36
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "DOC. I.D:"
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
         Left            =   4800
         TabIndex        =   35
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label10 
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
         Height          =   435
         Left            =   4800
         TabIndex        =   34
         Top             =   3720
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA:"
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
         Left            =   4800
         TabIndex        =   33
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TEL:"
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
         Left            =   3120
         TabIndex        =   32
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   31
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ACUDIENTE:"
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
         TabIndex        =   30
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R-H:"
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
         Left            =   7920
         TabIndex        =   29
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label Label4 
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
         Height          =   435
         Left            =   4800
         TabIndex        =   28
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label3 
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
         TabIndex        =   27
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         TabIndex        =   26
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. CARNET:"
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
         TabIndex        =   25
         Top             =   720
         Width           =   1185
      End
   End
End
Attribute VB_Name = "CONS_ALUM"
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
    Printer.CurrentY = 1.5
    Printer.CurrentX = 8
    Printer.Font.Size = 12
    Printer.Print "DATOS DEL ESTUDIANTE"
    Printer.CurrentY = 3
    Printer.CurrentX = 2
    Printer.Font.Size = 10
    Printer.Print ini.nombre
    Printer.CurrentY = 4
    Printer.CurrentX = 2
    Printer.Print "No.CARNET: " & Text1.Text
    Printer.CurrentY = 5
    Printer.CurrentX = 2
    Printer.Print "MATRICULA No: " & Text13.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "NOMBRES: ";
    Printer.Print Text2.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "APELLIDOS: ";
    Printer.Print Text3.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "DOCUMENTO DE I.D: ";
    Printer.Print Text11.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "FECHA NACIMIENTO: ";
    Printer.Print Text4.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "EDAD: ";
    Printer.Print Text15.Text & " AÑOS";
    Printer.CurrentX = 5
    Printer.Print "FACTOR R-H: ";
    Printer.Print Text5.Text;
    Printer.CurrentX = 8.3
    Printer.Print "SEXO: ";
    Printer.Print Text14.Text
    Printer.CurrentY = 5
    Printer.CurrentX = 11
    Printer.Print "ACUDIENTE: ";
    Printer.Print Text6.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 11
    Printer.Print "DIRECCION: ";
    Printer.Print Text7.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 11
    Printer.Print "EMAIL: ";
    Printer.Print Text23.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 11
    Printer.Print "TELEFONO: ";
    Printer.Print Text8.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 11
    Printer.Print "E.P.S: ";
    Printer.Print Text21.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 11
    Printer.Print "AÑO DE INGRESO: ";
    Printer.Print Text10.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 11
    Printer.Print "JORNADA: ";
    Printer.Print Text9.Text
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "GRADO: ";
    Printer.Print Text12.Text;
    Printer.CurrentX = 6.5
    Printer.Print "GRUPO: ";
    Printer.Print Text20.Text
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra toda la información de un alumno."
End Sub

