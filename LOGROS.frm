VERSION 5.00
Begin VB.Form LOGROS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OBSERVACIONES POR GRADO"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8895
   Icon            =   "LOGROS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "COPIA DE LOGROS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   8655
      Begin VB.OptionButton Option2 
         Caption         =   "EN DISKETTE"
         Height          =   375
         Left            =   1200
         TabIndex        =   35
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "EN DISCO DURO"
         Height          =   375
         Left            =   1200
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "LOGROS.frx":0ABA
         Left            =   5040
         List            =   "LOGROS.frx":0AEE
         TabIndex        =   33
         Text            =   "PREKINDER"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&COPIAR"
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
         Left            =   6960
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "LOGROS.frx":0B77
         Left            =   5040
         List            =   "LOGROS.frx":0B8A
         TabIndex        =   30
         Text            =   "SEGUNDO"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "LOGROS.frx":0BB8
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "PARA EL GRADO......."
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3360
         TabIndex        =   32
         Top             =   720
         Width           =   1620
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PARA EL PERIODO..."
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3360
         TabIndex        =   29
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CO&NSULTAR OBSERVACIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
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
      Height          =   315
      ItemData        =   "LOGROS.frx":0FFA
      Left            =   7320
      List            =   "LOGROS.frx":100D
      TabIndex        =   0
      Text            =   "PRIMERO"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   21
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&IMPRIMIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
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
      ForeColor       =   &H00C00000&
      Height          =   4335
      Left            =   2640
      TabIndex        =   12
      Top             =   480
      Width           =   6135
      Begin VB.CommandButton Command8 
         Caption         =   "O&PCIONES"
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
         TabIndex        =   9
         ToolTipText     =   "Corrige, Copia y pega observaciones"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&GUARDAR"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1800
         TabIndex        =   5
         Top             =   3120
         Width           =   4215
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "LOGROS.frx":103B
         Left            =   5160
         List            =   "LOGROS.frx":104B
         TabIndex        =   4
         Text            =   "L"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text3 
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
         Height          =   320
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text2 
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   2295
      End
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
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
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
         Height          =   315
         ItemData        =   "LOGROS.frx":105B
         Left            =   4200
         List            =   "LOGROS.frx":108F
         TabIndex        =   1
         Text            =   "PREKINDER"
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   315
         ItemData        =   "LOGROS.frx":1118
         Left            =   1440
         List            =   "LOGROS.frx":1128
         TabIndex        =   15
         Text            =   "UNICA"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   5775
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "NOMBRE:"
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
            Left            =   2400
            TabIndex        =   25
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "No.AREA:"
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
            TabIndex        =   24
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         TabIndex        =   26
         Top             =   2040
         Width           =   2295
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "INDICADOR:"
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
            Top             =   360
            Width           =   1110
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL..."
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
         Left            =   4680
         TabIndex        =   31
         Top             =   4080
         Width           =   795
      End
      Begin VB.Line Line2 
         DrawMode        =   1  'Blackness
         X1              =   240
         X2              =   3480
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         DrawMode        =   1  'Blackness
         X1              =   240
         X2              =   3480
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "OBSERVACION:"
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
         TabIndex        =   20
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO DE OBSERVACION..."
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
         TabIndex        =   19
         Top             =   2400
         Width           =   2610
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "GRADO..."
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
         Left            =   3000
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA..."
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
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   120
      Picture         =   "LOGROS.frx":1149
      ScaleHeight     =   4755
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO..."
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
      Left            =   6120
      TabIndex        =   22
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "OBSERVACIONES POR GRADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "LOGROS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
For i = 0 To Combo1.ListCount - 1
If Combo1.Text <> Combo1.List(i) Then
Combo1.Text = Combo1.List(0)
End If
Next i
End Sub

Private Sub Combo2_Change()
For i = 0 To Combo2.ListCount - 1
If Combo2.Text <> Combo2.List(i) Then
Combo2.Text = Combo2.List(0)
End If
Next i
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SetFocus
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub


Private Sub Combo4_Change()
For i = 0 To Combo4.ListCount - 1
If Combo4.Text <> Combo4.List(i) Then
Combo4.Text = Combo4.List(0)
End If
Next i
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo2.SetFocus
End If
End Sub

Private Sub Combo5_Change()
For i = 0 To Combo5.ListCount - 1
If Combo5.Text <> Combo5.List(i) Then
Combo5.Text = Combo5.List(0)
End If
Next i
End Sub

Private Sub Combo6_Change()
For i = 0 To Combo6.ListCount - 1
If Combo6.Text <> Combo6.List(i) Then
Combo6.Text = Combo6.List(0)
End If
Next i
End Sub

Private Sub Command1_Click()
Dim logru As logris
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DEL AREA Y DE CLICK EN OK", 64, "GUARDAR"
Text1.SetFocus
GoTo SALUD
End If
If Text4.Text = "" Then
MsgBox "ESCRIBA LA OBSERVACION", 64, "GUARDAR"
Text4.SetFocus
GoTo SALUD
End If
NAR = FreeFile
Open "c:\windows\datos\" & fl & ser & Val(Text1.Text) & lw & ".lgr" For Random As #NAR Len = Len(logru)
logru.indicador = Format(Combo3.Text, ">")
logru.observ = Text4.Text
Put #NAR, Text3.Text, logru
Close #NAR
Text5.Text = Text5.Text + 1
Text4.Text = ""
Text3.Text = Text3.Text + 1
Combo3.SetFocus
SALUD:
End Sub

Private Sub Command2_Click()
Dim mate As infomater
Dim argra As areagr
Dim logru As logris
If Text1.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DE AREA", 64, "OBSERVACIONES"
Text1.SetFocus
GoTo EX44
End If
cli = Val(Text1.Text)
NAR = FreeFile
Open "c:\windows\datos\materia.edu" For Random As #NAR Len = Len(mate)
que = 0
While Not EOF(NAR)
que = que + 1
Get #NAR, que, mate
Wend
Close #NAR
If ((cli > (que - 1)) Or (cli < 1)) Then
MsgBox "NO EXISTE EL AREA", 16, "OBSERVACIONES"
Text1.SetFocus
Text2.Text = ""
GoTo EX44
End If
cona = 0
Open "c:\windows\datos\areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
cona = cona + 1
Get #NAR, cona, argra
If (RTrim(argra.grado) = RTrim(Combo2.Text) And (argra.num_area = Val(Text1.Text))) Then
Close #NAR
GoTo intel
End If
Wend
Close #NAR
MsgBox "ESTA AREA NO ESTA CREADA PARA ESTE GRADO", 16, "OBSERVACIONES"
Text1.SetFocus
Text2.Text = ""
GoTo EX44
intel:
If RTrim(Combo4.Text) = "PRIMERO" Then
lw = 1
End If
If RTrim(Combo4.Text) = "SEGUNDO" Then
lw = 2
End If
If RTrim(Combo4.Text) = "TERCERO" Then
lw = 3
End If
If RTrim(Combo4.Text) = "CUARTO" Then
lw = 4
End If
If RTrim(Combo4.Text) = "FINAL" Then
lw = 5
End If
If Combo1.Text = "UNICA" Then
fl = "1"
End If
If Combo1.Text = "MAÑANA" Then
fl = "2"
End If
If Combo1.Text = "TARDE" Then
fl = "3"
End If
If Combo1.Text = "NOCHE" Then
fl = "4"
End If
ser = Left(Combo2.Text, 3)
CROA = 0
Open "c:\windows\datos\" & fl & ser & Val(Text1.Text) & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
CROA = CROA + 1
Get #NAR, CROA, logru
Wend
Close #NAR
Open "c:\windows\datos\materia.edu" For Random As #NAR Len = Len(mate)
Get #NAR, cli, mate
Close #NAR
Text5.Text = CROA - 1
Text2.Text = RTrim(mate.nom)
Text1.Locked = True
Text3.Text = ""
Text4.Text = ""
Text3.Text = CROA
Combo3.SetFocus
Frame1.Caption = Combo2.Text & " - " & "PERIODO: " & Combo4.Text & " - " & "AREA: " & Text2.Text
EX44:
End Sub
Private Sub Command3_Click()
Dim logru As logris
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DE AREA Y PRESIONE OK", 16, "CONSULTAR"
Text1.SetFocus
Exit Sub
End If
If Text5.Text = 0 Then
MsgBox "NO EXISTE INFORMACION", 16, "CONSULTAR"
Exit Sub
End If
NAR = FreeFile
Open "c:\windows\datos\" & fl & ser & Val(Text1.Text) & lw & ".lgr" For Random As #NAR Len = Len(logru)
For i = 1 To Text5.Text
Get #NAR, i, logru
CONS_OBSER.MATI11.Rows = i + 1
CONS_OBSER.MATI11.Row = i
CONS_OBSER.MATI11.Col = 0
CONS_OBSER.MATI11.Text = i
CONS_OBSER.MATI11.Col = 1
CONS_OBSER.MATI11.Text = logru.indicador
CONS_OBSER.MATI11.Col = 2
CONS_OBSER.MATI11.Text = logru.observ
Next i
Close #NAR
CONS_OBSER.Frame1.Caption = Frame1.Caption
CONS_OBSER.Show
End Sub

Private Sub Command4_Click()
Dim logru As logris
Dim ini As inicio
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DE AREA Y PRESIONE OK", 16, "IMPRIMIR"
Text1.SetFocus
GoTo CHAO
End If
If Text5.Text = 0 Then
MsgBox "NO EXISTE INFORMACION PARA IMPRIMIR", 16, "IMPRIMIR"
GoTo CHAO
End If
RESP = MsgBox("DESEA IMPRIMIR LAS OBSERVACIONES EXISTENTES?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
NAR = FreeFile
Open "c:\windows\datos\inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.telefono
Close #NAR
Printer.ScaleMode = 7
Printer.Font.Size = 9
Printer.CurrentY = 2
Printer.CurrentX = 1.5
Printer.Print "REPORTE DE OBSERVACIONES GRADO: " & Frame1.Caption
Printer.CurrentY = 3
Printer.CurrentX = 1.5
Printer.Print ini.nombre;
Printer.CurrentX = 17
Printer.Print "FECHA: " & Date
Printer.CurrentX = 1.5
Printer.Print "JORNADA: " & Combo1.Text;
Printer.CurrentX = 8
Printer.Print "COD AREA: " & Val(Text1.Text)
Printer.Print ""
Printer.CurrentX = 1.5
Printer.Print "CD";
Printer.CurrentX = 2.5
Printer.Print "IND";
Printer.CurrentX = 3.5
Printer.Print "OBSERVACION"
Printer.Print ""
L = 0
Open "c:\windows\datos\" & fl & ser & Val(Text1.Text) & lw & ".lgr" For Random As #NAR Len = Len(logru)
For i = 1 To Text5.Text
Get #NAR, i, logru
Printer.CurrentX = 1.5
Printer.Print i;
Printer.CurrentX = 2.5
Printer.Print logru.indicador;
X = 85
L1 = Left(logru.observ, X)
While Right(L1, 1) <> " "
X = X - 1
L1 = Left(L1, X)
Wend
Printer.CurrentX = 3.5
Printer.Print L1
L = L + 1
If (L Mod 52) = 0 Then
Printer.NewPage
Printer.CurrentY = 4.5
End If
y = Len(L1)
y = 150 - y
L2 = Right(logru.observ, y)
If RTrim(L2) <> "" Then
Printer.CurrentX = 3.5
Printer.Print L2
L = L + 1
If (L Mod 52) = 0 Then
Printer.NewPage
Printer.CurrentY = 4.5
End If
End If
Printer.Print ""
L = L + 1
If (L Mod 52) = 0 Then
Printer.NewPage
Printer.CurrentY = 4.5
End If
Next i
Close #NAR
Printer.EndDoc
End If
CHAO:
End Sub
Private Sub Command6_Click()
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DE AREA Y PRESIONE OK", 16, "BORRAR"
Text1.SetFocus
GoTo NOLO
End If
If Text5.Text = 0 Then
MsgBox "NO EXISTE INFORMACION PARA REALIZAR LA COPIA", 16, "ADVERTENCIA"
GoTo NOLO
End If
If RTrim(Combo4.Text) = "PRIMERO" Then
lw = 1
End If
If RTrim(Combo4.Text) = "SEGUNDO" Then
lw = 2
End If
If RTrim(Combo4.Text) = "TERCERO" Then
lw = 3
End If
If RTrim(Combo4.Text) = "CUARTO" Then
lw = 4
End If
If RTrim(Combo4.Text) = "FINAL" Then
lw = 5
End If
If RTrim(Combo6.Text) = "PRIMERO" Then
lw2 = 1
End If
If RTrim(Combo6.Text) = "SEGUNDO" Then
lw2 = 2
End If
If RTrim(Combo6.Text) = "TERCERO" Then
lw2 = 3
End If
If RTrim(Combo6.Text) = "CUARTO" Then
lw2 = 4
End If
If RTrim(Combo6.Text) = "FINAL" Then
lw2 = 5
End If
ser2 = Left(Combo5.Text, 3)
If Option1.Value = True Then
RESP = MsgBox("DESEA REALIZAR LA COPIA PARA EL PERIODO " & Combo6.Text & " DE " & Combo5.Text, vbYesNo + vbQuestion + vbDefaultButton1, "COPIAR LOGROS")
If RESP = vbYes Then
If Dir("c:\windows\datos\" & fl & ser & Val(Text1.Text) & lw & ".lgr") = "" Then
MsgBox "NO EXISTEN LOGROS DEL PERIODO " & Combo4.Text, 32, "COPIAR LOGROS"
Combo4.SetFocus
GoTo NOLO
End If
FileCopy "c:\windows\datos\" & fl & ser & Val(Text1.Text) & lw & ".lgr", "c:\windows\datos\" & fl & ser2 & Val(Text1.Text) & lw2 & ".lgr"
MsgBox "COPIA TERMINO CON EXITO", 64, "COPIA"
End If
End If
If Option2.Value = True Then
RESP = MsgBox("DESEA REALIZAR LA COPIA PARA EL PERIODO " & Combo6.Text & " DE " & Combo5.Text & "?", vbYesNo + vbQuestion + vbDefaultButton1, "COPIAR LOGROS")
If RESP = vbYes Then
On Error Resume Next
Err.Clear
MsgBox "INSERTE EL DISKETTE DESTINO Y PRESIONE ACEPTAR", 64, "COPIA"
If Dir("a:\datos\inicial.edu") = "" Then
MsgBox "DISKETTE DESTINO NO LO INSERTO EN LA UNIDAD A O NO CORRESPONDE AL SISTEMA, COPIA NO EXITOSA", 16, "ADVERTENCIA"
GoTo NOLO
End If
FileCopy "c:\windows\datos\" & fl & ser & Val(Text1.Text) & lw & ".lgr", "a:\datos\" & fl & ser2 & Val(Text1.Text) & lw2 & ".lgr"
If Err.Number = 70 Then
MsgBox "EL DISKETTE ESTA PROTEGIDO CONTRA ESCRITURA, NO SE REALIZO LA COPIA.", 16, "ADVERTENCIA"
GoTo NOLO
End If
If Err.Number = 61 Then
MsgBox "EL DISKETTE ESTA LLENO NO SE GUARDO LA INFORMACION", 16, "ADVERTENCIA"
GoTo NOLO
End If
MsgBox "COPIA TERMINO CON EXITO", 64, "COPIA"
End If
End If
NOLO:
End Sub

Private Sub Command8_Click()
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DE AREA Y PRESIONE OK", 16, "CONSULTAR"
Text1.SetFocus
Exit Sub
End If
i = 0
PASSW.Show 1
If i = 1 Then
COPEGA.Label2.Caption = "c:\windows\datos\" & fl & ser & Val(Text1.Text) & lw & ".lgr"
COPEGA.Label3.Caption = Text5.Text
COPEGA.Frame1.Caption = "JORNADA: " & Combo1.Text & " - GRADO: " & Combo2.Text & " - AREA: " & Text2.Text & " - PERIODO: " & Combo4.Text
Unload Me
COPEGA.Show
End If
End Sub

Private Sub Text1_DblClick()
Text1.Locked = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command2_Click
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONC42
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONC42:
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo3.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONC43
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONC43:
End Sub
Private Sub Form_Load()
Text1.MaxLength = 2
Text4.MaxLength = 150
Option1.Value = True
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
