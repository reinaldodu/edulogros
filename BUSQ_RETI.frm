VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BUSQ_RETI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar por: jornada, grado y año de retiro"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   Icon            =   "BUSQ_RETI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   4320
      Picture         =   "BUSQ_RETI.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Copiar la información que se muestra en pantalla"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   5520
      Picture         =   "BUSQ_RETI.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir la información que se muestra en pantalla"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   3120
      Picture         =   "BUSQ_RETI.frx":0A76
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ordenar los registros por apellidos ascendentemente"
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9255
      Begin MSFlexGridLib.MSFlexGrid MATI8 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   12632256
         GridColor       =   12582912
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   4320
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "BUSQ_RETI.frx":0B78
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "BUSQ_RETI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MATI8.Col = 0
MATI8.Sort = 5
End Sub

Private Sub Command2_Click()
'Dim ini As inicio
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    PAG = 1
    Printer.CurrentY = 1
    Printer.CurrentX = 1
    Printer.Font.Size = 10
    Printer.Print "ALUMNOS RETIRADOS " & Frame1.Caption
    Printer.CurrentY = 1.5
    Printer.CurrentX = 1
    Printer.Print ini.nombre;
    Printer.CurrentX = 19
    Printer.Print "Pág." & PAG
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print "APELLIDOS";
    Printer.CurrentX = 5
    Printer.Print "NOMBRES";
    Printer.CurrentX = 9
    Printer.Print "INGRESO";
    Printer.CurrentX = 11
    Printer.Print "DIRECCION";
    Printer.CurrentX = 17.5
    Printer.Print "TELEFONO"
    Printer.Print ""
    Printer.Font.Size = 8
    For gua = 1 To (MATI8.Rows - 1)
        Printer.CurrentX = 1
        Printer.Print MATI8.TextMatrix(gua, 0);
        Printer.CurrentX = 5
        Printer.Print MATI8.TextMatrix(gua, 1);
        Printer.CurrentX = 9
        Printer.Print MATI8.TextMatrix(gua, 4);
        Printer.CurrentX = 11
        Printer.Print MATI8.TextMatrix(gua, 2);
        Printer.CurrentX = 17.5
        Printer.Print MATI8.TextMatrix(gua, 3)
        If (gua Mod 67) = 0 Then
            Printer.NewPage
            PAG = PAG + 1
            Printer.CurrentY = 1
            Printer.CurrentX = 1
            Printer.Font.Size = 10
            Printer.Print "ALUMNOS RETIRADOS " & Frame1.Caption
            Printer.CurrentY = 1.5
            Printer.CurrentX = 1
            Printer.Print ini.nombre;
            Printer.CurrentX = 19
            Printer.Print "Pág." & PAG
            Printer.Print ""
            Printer.CurrentX = 1
            Printer.Print "APELLIDOS";
            Printer.CurrentX = 5
            Printer.Print "NOMBRES";
            Printer.CurrentX = 9
            Printer.Print "INGRESO";
            Printer.CurrentX = 11
            Printer.Print "DIRECCION";
            Printer.CurrentX = 17.5
            Printer.Print "TELEFONO"
            Printer.Print ""
            Printer.Font.Size = 8
        End If
    Next gua
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print "TOTAL RETIRADOS..." & MATI8.Rows - 1
    Printer.EndDoc
    Screen.MousePointer = 0
End If
End Sub

Private Sub Command3_Click()
Clipboard.Clear
cop = ""
cop = "ALUMNOS RETIRADOS " & Frame1.Caption & vbCrLf & vbCrLf
For X = 1 To (MATI8.Rows - 1)
        ape = RTrim(MATI8.TextMatrix(X, 0))
        nom = RTrim(MATI8.TextMatrix(X, 1))
        If X < 10 Then
           cop = cop + LTrim(ape & " " & nom) & vbCrLf
        Else
           cop = cop + LTrim(ape & " " & nom) & vbCrLf
        End If
Next X
Clipboard.SetText cop
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra la información de alumnos retirados, de acuerdo a la jornada, grado y año de retiro."
End Sub

Private Sub Form_Load()
Frame1.Caption = "JORNADA: " & RETIRADOS.Combo1.Text & "   GRADO: " & RETIRADOS.Combo2.Text & "   AÑO DE RETIRO: " & RETIRADOS.Combo3.Text
MATI8.Row = 0
MATI8.Col = 0
MATI8.ColWidth(0) = 2400
MATI8.CellForeColor = RGB(255, 255, 255)
MATI8.CellBackColor = RGB(0, 0, 150)
MATI8.Text = "APELLIDOS"
MATI8.Col = 1
MATI8.ColWidth(1) = 2400
MATI8.CellForeColor = RGB(255, 255, 255)
MATI8.CellBackColor = RGB(0, 0, 150)
MATI8.Text = "NOMBRES"
MATI8.Col = 2
MATI8.ColWidth(2) = 4000
MATI8.CellForeColor = RGB(255, 255, 255)
MATI8.CellBackColor = RGB(0, 0, 150)
MATI8.Text = "DIRECCION"
MATI8.Col = 3
MATI8.ColWidth(3) = 1000
MATI8.CellForeColor = RGB(255, 255, 255)
MATI8.CellBackColor = RGB(0, 0, 150)
MATI8.Text = "TELEFONO"
MATI8.Col = 4
MATI8.CellForeColor = RGB(255, 255, 255)
MATI8.CellBackColor = RGB(0, 0, 150)
MATI8.Text = "INGRESO"
End Sub
