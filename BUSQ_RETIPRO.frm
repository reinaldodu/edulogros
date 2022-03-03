VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BUSQ_RETIPRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profesores retirados - Buscar por año de retiro"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   Icon            =   "BUSQ_RETIPRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   4320
      Picture         =   "BUSQ_RETIPRO.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Copiar la información que se muestra en pantalla"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   3120
      Picture         =   "BUSQ_RETIPRO.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ordenar los registros por nombres ascendentemente"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   5520
      Picture         =   "BUSQ_RETIPRO.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir la información que se muestra en pantalla"
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
      TabIndex        =   5
      Top             =   600
      Width           =   9255
      Begin MSFlexGridLib.MSFlexGrid MATI30 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   12632256
         GridColor       =   12582912
      End
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
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   120
      Picture         =   "BUSQ_RETIPRO.frx":0B78
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "BUSQ_RETIPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "BUSQUEDA POR AÑO DE RETIRO")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    PAG = 1
    noyu = 0
    Printer.ScaleMode = 7
    Printer.CurrentY = 1
    Printer.CurrentX = 6
    Printer.Font.Size = 10
    Printer.Print "PROFESORES RETIRADOS " & "(" & Frame1.Caption & ")"
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print ini.nombre;
    Printer.CurrentX = 19
    Printer.Print "Pág." & PAG
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print "NOMBRES Y APELLIDOS";
    Printer.CurrentX = 8
    Printer.Print "CEDULA";
    Printer.CurrentX = 10.5
    Printer.Print "DIRECCION";
    Printer.CurrentX = 18
    Printer.Print "TELEFONO"
    Printer.Print ""
    For q = 1 To (MATI30.Rows - 1)
        Printer.CurrentX = 1
        Printer.Print MATI30.TextMatrix(q, 0);
        Printer.CurrentX = 8
        Printer.Print MATI30.TextMatrix(q, 1);
        Printer.CurrentX = 10.5
        Printer.Print MATI30.TextMatrix(q, 2);
        Printer.CurrentX = 18
        Printer.Print MATI30.TextMatrix(q, 3)
        noyu = noyu + 1
        If (noyu Mod 53) = 0 Then
            Printer.NewPage
            PAG = PAG + 1
            Printer.CurrentY = 1
            Printer.CurrentX = 6
            Printer.Font.Size = 10
            Printer.Print "PROFESORES RETIRADOS " & "(" & Frame1.Caption & ")"
            Printer.Print ""
            Printer.Print ""
            Printer.CurrentX = 1
            Printer.Print ini.nombre;
            Printer.CurrentX = 19
            Printer.Print "Pág." & PAG
            Printer.Print ""
            Printer.CurrentX = 1
            Printer.Print "NOMBRES Y APELLIDOS";
            Printer.CurrentX = 8
            Printer.Print "CEDULA";
            Printer.CurrentX = 10.5
            Printer.Print "DIRECCION";
            Printer.CurrentX = 18
            Printer.Print "TELEFONO"
            Printer.Print ""
        End If
    Next q
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print "TOTAL RETIRADOS..." & MATI30.Rows - 1
    Printer.EndDoc
    Printer.Font.Size = 8
    Screen.MousePointer = 0
End If
End Sub

Private Sub Command2_Click()
MATI30.Col = 0
MATI30.Sort = 5
End Sub

Private Sub Command3_Click()
Clipboard.Clear
cop = ""
cop = "PROFESORES RETIRADOS " & "(" & Frame1.Caption & ")" & vbCrLf & vbCrLf
For X = 1 To (MATI30.Rows - 1)
        ape = RTrim(MATI30.TextMatrix(X, 0))
        If X < 10 Then
           cop = cop + LTrim(ape) & vbCrLf
        Else
           cop = cop + LTrim(ape) & vbCrLf
        End If
Next X
Clipboard.SetText cop
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra la información de profesores retirados en un año determinado."
End Sub

Private Sub Form_Load()
MATI30.Row = 0
MATI30.Col = 0
MATI30.CellFontBold = True
MATI30.CellForeColor = RGB(0, 0, 255)
MATI30.ColWidth(0) = 3500
MATI30.Text = "NOMBRES Y APELLIDOS"
MATI30.Col = 1
MATI30.CellFontBold = True
MATI30.CellForeColor = RGB(0, 0, 255)
MATI30.ColWidth(1) = 1000
MATI30.Text = "CEDULA"
MATI30.Col = 2
MATI30.CellFontBold = True
MATI30.CellForeColor = RGB(0, 0, 255)
MATI30.ColWidth(2) = 4000
MATI30.Text = "DIRECCION"
MATI30.Col = 3
MATI30.CellFontBold = True
MATI30.CellForeColor = RGB(0, 0, 255)
MATI30.ColWidth(3) = 1100
MATI30.Text = "TELEFONO"
End Sub
