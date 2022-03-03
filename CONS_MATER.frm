VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONS_MATER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de materias"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6390
   Icon            =   "CONS_MATER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.TextBox Text1 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   1920
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid MATI6 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   1
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total materias existentes..."
         Height          =   195
         Left            =   3480
         TabIndex        =   3
         Top             =   2040
         Width           =   1875
      End
   End
End
Attribute VB_Name = "CONS_MATER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim mate As infomater
'Dim ini As inicio
If Val(Text1.Text) = 0 Then
    MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 32, "IMPRIMIR"
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR LA LISTA DE AREAS?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR AREAS")
If RESP = vbYes Then
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
Printer.ScaleMode = 7
Printer.CurrentY = 1
Printer.CurrentX = 7.5
Printer.Font.Size = 10
Printer.Print "LISTA DE AREAS EXISTENTES"
Printer.Print ""
Printer.Print ""
Printer.CurrentX = 1
Printer.Print ini.nombre
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "No.";
Printer.CurrentX = 2
Printer.Print "AREA"
Printer.Print ""
For maa = 1 To (MATI6.Rows - 1)
    Printer.CurrentX = 1
    Printer.Print MATI6.TextMatrix(maa, 0);
    Printer.CurrentX = 2
    Printer.Print MATI6.TextMatrix(maa, 1)
Next maa
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra las áreas existentes."
End Sub

Private Sub Form_Load()
'Dim mate As infomater
MATI6.Rows = 1
L = 1
MATI6.Row = 0
MATI6.Col = 0
MATI6.ColWidth(0) = 400
MATI6.CellForeColor = RGB(255, 255, 255)
MATI6.CellBackColor = RGB(0, 0, 150)
MATI6.Text = "No."
MATI6.Col = 1
MATI6.ColWidth(1) = 5000
MATI6.CellForeColor = RGB(255, 255, 255)
MATI6.CellBackColor = RGB(0, 0, 150)
MATI6.Text = "NOMBRE"
If Dir(Ruta & "materia.edu") <> "" Then
    h = 0
    NAR = FreeFile
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    que = 0
    While Not EOF(NAR)
        que = que + 1
        Get #NAR, que, mate
        If RTrim(mate.nom) = "" Then
            h = h + 1
        End If
        If mate.num = 0 Then
            GoTo FING
        End If
        MATI6.Rows = MATI6.Rows + 1
        MATI6.TextMatrix(L, 0) = mate.num
        MATI6.TextMatrix(L, 1) = mate.nom
        L = L + 1
    Wend
FING:
    Close #NAR
    Text1.Text = (que - 1) - h
End If
End Sub
