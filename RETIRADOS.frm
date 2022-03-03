VERSION 5.00
Begin VB.Form RETIRADOS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base de datos de alumnos retirados"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
   Icon            =   "RETIRADOS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&ELIMINAR RETIRADO"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Elimina el registro que se muestra en pantalla de la base de datos de alumnos retirados"
      Top             =   5160
      Width           =   2175
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
      ForeColor       =   &H0000FFFF&
      Height          =   320
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   5040
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "BUSCAR POR:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   2640
      TabIndex        =   25
      Top             =   3120
      Width           =   6255
      Begin VB.Frame Frame4 
         Caption         =   "JORNADA, GRADO Y AÑO DE RETIRO"
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   3000
         TabIndex        =   29
         Top             =   240
         Width           =   3135
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   960
            TabIndex        =   5
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Ok"
            Height          =   315
            Left            =   2400
            TabIndex        =   6
            Top             =   960
            Width           =   600
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "RETIRADOS.frx":0442
            Left            =   960
            List            =   "RETIRADOS.frx":047C
            TabIndex        =   4
            Text            =   "PREKINDER"
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "RETIRADOS.frx":051B
            Left            =   960
            List            =   "RETIRADOS.frx":052B
            TabIndex        =   3
            Text            =   "UNICA"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "AÑO:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "GRADO:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "JORNADA:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "NOMBRES Y APELLIDOS"
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2775
         Begin VB.CommandButton Command2 
            Caption         =   "O&k"
            Height          =   255
            Left            =   1080
            TabIndex        =   2
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox Text10 
            Height          =   320
            Left            =   1080
            TabIndex        =   1
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox Text9 
            Height          =   320
            Left            =   1080
            TabIndex        =   0
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "APELLIDOS:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   930
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "NOMBRES:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&IMPRIMIR"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   "Imprime los datos personales mostrados en pantalla"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   120
      Picture         =   "RETIRADOS.frx":054C
      ScaleHeight     =   4755
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATOS PERSONALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   2640
      TabIndex        =   9
      Top             =   0
      Width           =   6255
      Begin VB.TextBox Text8 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "AÑO DE RETIRO:"
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
         Left            =   3120
         TabIndex        =   26
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label6 
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
         Left            =   3120
         TabIndex        =   16
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   3120
         TabIndex        =   15
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label4 
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
         Left            =   3120
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1050
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
         Left            =   120
         TabIndex        =   12
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
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   990
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL RETIRADOS..."
      Height          =   195
      Left            =   6480
      TabIndex        =   35
      Top             =   5160
      Width           =   1650
   End
End
Attribute VB_Name = "RETIRADOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text <> Combo1.List(0) Then
    Combo1.Text = Combo1.List(0)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo3.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command3_Click
End If
End Sub

Private Sub Command1_Click()
'Dim ini As inicio
If Text2.Text = "" Then
    MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 64, "IMPRIMIR"
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR LA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
Close #NAR
Printer.ScaleMode = 7
Printer.CurrentY = 1
Printer.CurrentX = 8
Printer.Font.Size = 12
Printer.Print "ESTUDIANTE RETIRADO"
Printer.Font.Size = 10
Printer.Print ""
Printer.Print ""
Printer.CurrentX = 2
Printer.Print ini.nombre
Printer.CurrentY = 4
Printer.CurrentX = 2
Printer.Print "APELLIDOS: " & Text2.Text;
Printer.CurrentX = 9
Printer.Print "NOMBRES: " & Text1.Text
Printer.Print ""
Printer.CurrentX = 2
Printer.Print "TELEFONO: " & Text3.Text;
Printer.CurrentX = 9
Printer.Print "DIRECCION: " & Text5.Text
Printer.Print ""
Printer.CurrentX = 2
Printer.Print "JORNADA: " & Text6.Text;
Printer.CurrentX = 9
Printer.Print "GRADO: " & Text4.Text
Printer.Print ""
Printer.CurrentX = 2
Printer.Print "AÑO DE INGRESO: " & Text7.Text;
Printer.CurrentX = 9
Printer.Print "AÑO DE RETIRO: " & Text8.Text
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Command2_Click()
'Dim retiros As retiro
If Text9.Text = "" Then
    MsgBox "ESCRIBA LOS NOMBRES", 16, "ADVERTENCIA"
    Text9.SetFocus
    Exit Sub
End If
If Text10.Text = "" Then
    MsgBox "ESCRIBA LOS APELLIDOS", 16, "ADVERTENCIA"
    Text10.SetFocus
    Exit Sub
End If
Text9.Text = Format(Text9.Text, ">")
Text10.Text = Format(Text10.Text, ">")
cruz = 0
NAR = FreeFile
Open Ruta & "retialu.edu" For Random As #NAR Len = Len(retiros)
While Not EOF(NAR)
cruz = cruz + 1
Get #NAR, cruz, retiros
If (((RTrim(retiros.nombres) = RTrim(Text9.Text)) And (RTrim(retiros.apellidos) = RTrim(Text10.Text)))) Then
Text1.Text = retiros.nombres
Text2.Text = retiros.apellidos
Text5.Text = retiros.direccion
Text3.Text = retiros.Telefono
Text6.Text = retiros.jornada
Text7.Text = retiros.año_ingreso
Text4.Text = retiros.grado
Text8.Text = retiros.año_retiro
Close #NAR
Exit Sub
End If
Wend
Close #NAR
MsgBox "NO SE ENCONTRO REGISTRO", 48, "BUSQUEDA"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.SetFocus
End Sub
Private Sub Command3_Click()
'Dim retiros As retiro
falta = 1
ja = 0
NAR = FreeFile
Open Ruta & "retialu.edu" For Random As #NAR Len = Len(retiros)
While Not EOF(NAR)
    ja = ja + 1
    Get #NAR, ja, retiros
    If ((RTrim(retiros.jornada) = RTrim(Combo1.Text)) And (RTrim(retiros.grado) = RTrim(Combo2.Text)) And (retiros.año_retiro = Combo3.Text)) Then
        BUSQ_RETI.MATI8.Rows = falta + 1
        BUSQ_RETI.MATI8.TextMatrix(falta, 0) = RTrim(retiros.apellidos)
        BUSQ_RETI.MATI8.TextMatrix(falta, 1) = RTrim(retiros.nombres)
        BUSQ_RETI.MATI8.TextMatrix(falta, 2) = RTrim(retiros.direccion)
        BUSQ_RETI.MATI8.TextMatrix(falta, 3) = RTrim(retiros.Telefono)
        BUSQ_RETI.MATI8.TextMatrix(falta, 4) = RTrim(retiros.año_ingreso)
        falta = falta + 1
    End If
Wend
Close #NAR
If falta = 1 Then
    MsgBox "NO SE ENCONTRARON REGISTROS", 16, "BUSQUEDA"
    Combo3.SetFocus
    Exit Sub
End If
BUSQ_RETI.Label1.Caption = "REGISTROS ENCONTRADOS = " & (falta - 1)
BUSQ_RETI.Show
End Sub

Private Sub Command4_Click()
'Dim retiros As retiro
I = 0
PASSW.Show 1
If I = 1 Then
If Text1.Text = "" Then
    MsgBox "ESCRIBA LOS NOMBRES Y APELLIDOS Y DE CLICK EN OK PARA PODER ELIMINAR UN ESTUDIANTE DE LA BASE DE DATOS DE RETIRADOS", 48, "ELIMINAR RETIRADO"
    Text9.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA ELIMINAR ESTE ESTUDIANTE?", vbYesNo + vbQuestion + vbDefaultButton2, "ELIMINAR RETIRADO")
If RESP = vbYes Then
NAR = FreeFile
Open Ruta & "retialu.edu" For Random As #NAR Len = Len(retiros)
retiros.año_ingreso = ""
retiros.año_retiro = ""
retiros.apellidos = ""
retiros.direccion = ""
retiros.grado = ""
retiros.jornada = ""
retiros.nombres = ""
retiros.Telefono = ""
Put #NAR, cruz, retiros
Close #NAR
Open Ruta & "conelire.edu" For Input As #NAR
Input #NAR, z
Close #NAR
z = z + 1
Open Ruta & "conelire.edu" For Output As #NAR
Print #NAR, z
Close #NAR
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text12.Text = Text12.Text - 1
End If
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Base de datos de alumnos retirados."
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command2_Click
End If
End Sub

Private Sub Form_Load()
For I = 1998 To 2100
Combo3.AddItem I
Next I
Combo3.Text = Combo3.List(0)
NAR = FreeFile
Open Ruta & "contreti.edu" For Input As #NAR
Input #NAR, zi
Close #NAR
Open Ruta & "conelire.edu" For Input As #NAR
Input #NAR, z
Close #NAR
Text9.MaxLength = 20
Text10.MaxLength = 20
Text12.Text = (zi - 1) - z
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text10.SetFocus
End If
End Sub
