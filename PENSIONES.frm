VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PENSIONES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de pensiones"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "PENSIONES.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   635
      ForeColor       =   12582912
      MouseIcon       =   "PENSIONES.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PAZ Y SALVO"
      TabPicture(0)   =   "PENSIONES.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PAGOS Y PENDIENTES"
      TabPicture(1)   =   "PENSIONES.frx":0640
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame6 
         Caption         =   "2. PENSIONES PENDIENTES POR JORNADA Y GRADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74880
         TabIndex        =   30
         Top             =   3240
         Width           =   6135
         Begin VB.Frame Frame8 
            Height          =   1935
            Left            =   4200
            TabIndex        =   37
            Top             =   240
            Width           =   1815
            Begin VB.CommandButton Command7 
               Caption         =   "IMPRIMIR"
               Height          =   375
               Left            =   240
               TabIndex        =   40
               Top             =   1320
               Width           =   1335
            End
            Begin VB.CommandButton Command6 
               Caption         =   "ACEPTAR"
               Height          =   375
               Left            =   240
               TabIndex        =   39
               Top             =   720
               Width           =   1335
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
               ForeColor       =   &H0000FFFF&
               Height          =   315
               ItemData        =   "PENSIONES.frx":095A
               Left            =   240
               List            =   "PENSIONES.frx":096A
               TabIndex        =   38
               Text            =   "UNICA"
               Top             =   240
               Width           =   1335
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MATI99 
            Height          =   1815
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   3201
            _Version        =   393216
            Rows            =   224
            Cols            =   3
            BackColorBkg    =   -2147483633
            GridColor       =   12582912
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "1. TOTAL PAGOS POR JORNADA Y GRADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   29
         Top             =   720
         Width           =   6135
         Begin VB.Frame Frame7 
            Height          =   1815
            Left            =   4200
            TabIndex        =   32
            Top             =   240
            Width           =   1815
            Begin VB.CommandButton Command5 
               Caption         =   "IMPRIMIR"
               Height          =   375
               Left            =   240
               TabIndex        =   35
               Top             =   1200
               Width           =   1335
            End
            Begin VB.CommandButton Command4 
               Caption         =   "ACEPTAR"
               Height          =   375
               Left            =   240
               TabIndex        =   34
               Top             =   720
               Width           =   1335
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
               ItemData        =   "PENSIONES.frx":098B
               Left            =   240
               List            =   "PENSIONES.frx":099B
               TabIndex        =   33
               Text            =   "UNICA"
               Top             =   240
               Width           =   1335
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MATI88 
            Height          =   1695
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   2990
            _Version        =   393216
            Rows            =   224
            Cols            =   3
            BackColorBkg    =   -2147483633
            GridColor       =   12582912
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5295
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   5895
         Begin VB.Frame Frame2 
            Height          =   3375
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   5415
            Begin VB.CommandButton Command1 
               Caption         =   "&Aceptar"
               Height          =   375
               Left            =   240
               TabIndex        =   4
               Top             =   2760
               Width           =   1335
            End
            Begin VB.TextBox Text6 
               Height          =   285
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox Text5 
               Height          =   285
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox Text3 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   1320
               Width           =   1935
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   360
               Width           =   1935
            End
            Begin VB.Frame Frame4 
               Caption         =   "CARNET"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   700
               Left            =   3240
               TabIndex        =   19
               Top             =   1920
               Width           =   2055
               Begin VB.TextBox Text8 
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
                  Height          =   360
                  Left            =   480
                  TabIndex        =   1
                  Top             =   240
                  Width           =   735
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "Ok"
                  Height          =   360
                  Left            =   1320
                  TabIndex        =   2
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "No."
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
                  TabIndex        =   20
                  Top             =   360
                  Width           =   315
               End
            End
            Begin VB.Frame Frame3 
               Height          =   700
               Left            =   240
               TabIndex        =   17
               Top             =   1920
               Width           =   2895
               Begin VB.TextBox Text7 
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
                  ForeColor       =   &H00FFFFFF&
                  Height          =   360
                  Left            =   960
                  TabIndex        =   3
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "VALOR $"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   18
                  Top             =   360
                  Width           =   795
               End
            End
            Begin VB.CommandButton Command2 
               Caption         =   "&Imprimir"
               Height          =   375
               Left            =   3960
               TabIndex        =   5
               Top             =   2760
               Width           =   1335
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "DOC. ID:"
               Height          =   195
               Left            =   3240
               TabIndex        =   26
               Top             =   1440
               Width           =   645
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "JORNADA:"
               Height          =   195
               Left            =   3240
               TabIndex        =   25
               Top             =   480
               Width           =   810
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "GRUPO:"
               Height          =   195
               Left            =   240
               TabIndex        =   24
               Top             =   1440
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "GRADO:"
               Height          =   195
               Left            =   3240
               TabIndex        =   23
               Top             =   960
               Width           =   630
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "APELLIDOS:"
               Height          =   195
               Left            =   240
               TabIndex        =   22
               Top             =   960
               Width           =   930
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "NOMBRES:"
               Height          =   195
               Left            =   240
               TabIndex        =   21
               Top             =   480
               Width           =   855
            End
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
            ItemData        =   "PENSIONES.frx":09BC
            Left            =   840
            List            =   "PENSIONES.frx":09E4
            TabIndex        =   0
            Text            =   "ENERO"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox Text9 
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
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid MATI77 
            Height          =   900
            Left            =   240
            TabIndex        =   14
            Top             =   4200
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   1588
            _Version        =   393216
            Cols            =   14
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "FECHA:"
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
            TabIndex        =   28
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "MES:"
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
            Width           =   465
         End
      End
      Begin VB.Label Label11 
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
         Left            =   -73320
         TabIndex        =   41
         Top             =   5640
         Width           =   75
      End
   End
End
Attribute VB_Name = "PENSIONES"
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
Text8.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Command1_Click()
'Dim pens(1 To 12) As Currency
If Text7.Text = "." Then Text7.Text = ""
If Text1.Text = "" Then
    MsgBox "ESCRIBA PRIMERO EL No. DE CARNET", 64, "CONTROL DE PENSIONES"
    Text8.SetFocus
    Exit Sub
End If
If Text7.Text = "" Then
    MsgBox "ESCRIBA PRIMERO EL VALOR", 64, "CONTROL DE PENSIONES"
    Text7.SetFocus
    Exit Sub
End If
If Text5.Text = "SIN GRADO" Then
    MsgBox "NO SE PUEDE CREAR PAZ Y SALVO DE ESTUDIANTE SIN GRADO", 32, "CREAR PAZ Y SALVO"
    Exit Sub
End If
RESP = MsgBox("DESEA CREAR ESTE PAZ Y SALVO?", vbYesNo + vbQuestion + vbDefaultButton1, "CONTROL DE PENSIONES")
If RESP = vbYes Then
    If Combo1.Text = "ENERO" Then
        rt = 1
    End If
    If Combo1.Text = "FEBRERO" Then
        rt = 2
    End If
    If Combo1.Text = "MARZO" Then
        rt = 3
    End If
    If Combo1.Text = "ABRIL" Then
        rt = 4
    End If
    If Combo1.Text = "MAYO" Then
        rt = 5
    End If
    If Combo1.Text = "JUNIO" Then
        rt = 6
    End If
    If Combo1.Text = "JULIO" Then
        rt = 7
    End If
    If Combo1.Text = "AGOSTO" Then
        rt = 8
    End If
    If Combo1.Text = "SEPTIEMBRE" Then
        rt = 9
    End If
    If Combo1.Text = "OCTUBRE" Then
        rt = 10
    End If
    If Combo1.Text = "NOVIEMBRE" Then
        rt = 11
    End If
    If Combo1.Text = "DICIEMBRE" Then
        rt = 12
    End If
    For J = 1 To 12
        If J = rt Then
            pens(J) = Text7.Text
            GoTo suns
        End If
        pens(J) = MATI77.TextMatrix(1, J)
suns:
    Next J
    NAR = FreeFile
    Open Ruta & "pensi.edu" For Random As #NAR Len = 96
    Put #NAR, h, pens
    Close #NAR
    MATI77.TextMatrix(1, rt) = Text7.Text
    trt = 0
    For Y = 1 To 12
        MATI77.Col = Y
        MATI77.CellForeColor = RGB(255, 255, 255)
        MATI77.CellBackColor = RGB(0, 0, 150)
        MATI77.Text = Format(pens(Y), "###,###,###.00")
        trt = trt + pens(Y)
    Next Y
    MATI77.TextMatrix(1, 13) = Format(trt, "###,###,###.00")
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
MsgBox "ESCRIBA PRIMERO EL No. DE CARNET", 64, "CONTROL DE PENSIONES"
Text8.SetFocus
Exit Sub
End If
If Combo1.Text = "ENERO" Then
rt = 1
End If
If Combo1.Text = "FEBRERO" Then
rt = 2
End If
If Combo1.Text = "MARZO" Then
rt = 3
End If
If Combo1.Text = "ABRIL" Then
rt = 4
End If
If Combo1.Text = "MAYO" Then
rt = 5
End If
If Combo1.Text = "JUNIO" Then
rt = 6
End If
If Combo1.Text = "JULIO" Then
rt = 7
End If
If Combo1.Text = "AGOSTO" Then
rt = 8
End If
If Combo1.Text = "SEPTIEMBRE" Then
rt = 9
End If
If Combo1.Text = "OCTUBRE" Then
rt = 10
End If
If Combo1.Text = "NOVIEMBRE" Then
rt = 11
End If
If Combo1.Text = "DICIEMBRE" Then
rt = 12
End If
MATI77.Row = 1
MATI77.Col = rt
If MATI77.Text = 0 Then
MsgBox "DEBE CREAR PRIMERO EL PAZ Y SALVO DEL MES DE " & Combo1.Text, 64, "IMPRIMIR"
Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTE PAZ Y SALVO?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR PAZ Y SALVO")
If RESP = vbYes Then
Printer.ScaleMode = 7
Printer.CurrentY = 1
Printer.CurrentX = 6
Printer.Font.Size = 14
Printer.Print "PAZ Y SALVO MES DE " & Combo1.Text
Printer.CurrentY = 2
Printer.CurrentX = 17
Printer.Font.Size = 8
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
Printer.CurrentY = 3
Printer.CurrentX = 2
Printer.Font.Size = 11
Printer.Print "NOMBRES: " & Text1.Text;
Printer.CurrentX = 11
Printer.Print "JORNADA: " & Text4.Text
Printer.Print ""
Printer.CurrentX = 2
Printer.Print "APELLIDOS: " & Text2.Text;
Printer.CurrentX = 11
Printer.Print "GRADO: " & Text5.Text
Printer.Print ""
Printer.CurrentX = 2
Printer.Print "DOC.ID: " & Text6.Text;
Printer.CurrentX = 11
Printer.Print "GRUPO: " & Text3.Text
Printer.CurrentY = 8
Printer.CurrentX = 2
Printer.Font.Size = 9
Printer.Print "FIRMA AUTORIZADA Y SELLO";
Printer.CurrentX = 11
Printer.Font.Size = 14
Printer.Print "VALOR...$" & MATI77.Text
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Command3_Click()
'Dim alumno As maestroalum
'Dim pens(1 To 12) As Currency
'Dim aluper As pertgrup
If Text8.Text = "" Then
   MsgBox "ESCRIBA UN NUMERO DE CARNET", 64, "CONTROL DE PENSIONES"
   Text8.SetFocus
   Exit Sub
End If
If Val(Text8.Text) > 32000 Then
   MsgBox "No. DE CARNET INVALIDO", 64, "CONTROL DE PENSIONES"
   Text8.SetFocus
   Exit Sub
End If
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
h = Val(Text8.Text)
If ((h > I - 1) Or (h < 1)) Then
   MsgBox "REGISTRO NO EXISTE", 32
   Text1.Text = ""
   Text2.Text = ""
   Text4.Text = ""
   Text5.Text = ""
   Text6.Text = ""
   Text8.SetFocus
   Text8.Text = ""
   Exit Sub
End If
If Dir(Ruta & "pensi.edu") = "" Then
    Open Ruta & "pensi.edu" For Random As #NAR Len = 96
    For J = 1 To (I - 1)
        For Y = 1 To 12
            pens(Y) = 0
        Next Y
        Put #NAR, J, pens
    Next J
    Close #NAR
    Command1.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
End If
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
   MsgBox "REGISTRO NO EXISTE", 32
   Text8.SetFocus
   Exit Sub
End If
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
Get #NAR, h, aluper
Close #NAR
Text1.Text = RTrim(alumno.nombres)
Text2.Text = RTrim(alumno.apellidos)
Text3.Text = RTrim(aluper.grupo)
Text4.Text = RTrim(alumno.jornada)
Text5.Text = RTrim(alumno.grado)
Text6.Text = RTrim(alumno.documento)
Open Ruta & "pensi.edu" For Random As #NAR Len = 96
Get #NAR, h, pens
Close #NAR
MATI77.Row = 1
trt = 0
For rt = 1 To 12
    MATI77.Col = rt
    MATI77.CellForeColor = RGB(255, 255, 255)
    MATI77.CellBackColor = RGB(0, 0, 150)
    MATI77.Text = Format(pens(rt), "###,###,###.00")
    trt = trt + pens(rt)
Next rt
MATI77.TextMatrix(1, 13) = Format(trt, "###,###,###.00")
Text7.SetFocus
End Sub

Private Sub Command4_Click()
'Dim alumno As maestroalum
'Dim pens(1 To 12) As Currency
Label11.Caption = "ESPERE UN MOMENTO POR FAVOR..."
Screen.MousePointer = 11
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
Y = 1
tgp = 0
For k = 1 To 16
    spa = 0
    NAR = FreeFile
    Open Ruta & "pensi.edu" For Random As #NAR Len = 96
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    For J = 1 To 12
        h = 1
        tpm = 0
        While h < I
            Get #NAR, h, alumno
            Get #(NAR - 1), h, pens
            If ((RTrim(alumno.jornada) = Combo2.Text) And (RTrim(alumno.grado) = MATI88.TextMatrix(Y, 1))) Then
                tpm = tpm + pens(J)
            End If
            h = h + 1
        Wend
        MATI88.TextMatrix(Y, 2) = Format(tpm, "###,###,###.00")
        Y = Y + 1
        spa = spa + tpm
    Next J
    Close #NAR
    Close #(NAR - 1)
    MATI88.TextMatrix(Y, 0) = "SUBTOTAL..."
    MATI88.TextMatrix(Y, 2) = Format(spa, "###,###,###.00")
    Y = (k * 13) + 1
    tgp = tgp + spa
Next k
MATI88.TextMatrix(223, 2) = Format(tgp, "###,###,###.00")
For k = 1 To 12
    tgp = 0
    Y = k
    For J = 1 To 16
        tgp = tgp + MATI88.TextMatrix(Y, 2)
        Y = (J * 13) + k
    Next J
    Y = Y + 2
    MATI88.TextMatrix(Y, 2) = Format(tgp, "###,###,###.00")
Next k
Label11.Caption = ""
Screen.MousePointer = 0
End Sub

Private Sub Command5_Click()
MATI88.Row = 1
MATI88.Col = 2
If MATI88.Text = "" Then
    MsgBox "PRESIONE PRIMERO ACEPTAR", 64, "IMPRIMIR"
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR TOTAL DE PAGOS?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR TOTAL PAGOS")
If RESP = vbYes Then
Printer.ScaleMode = 7
Printer.CurrentY = 1
Printer.CurrentX = 5
Printer.Print "TOTAL PAGOS DE PENSIONES JORNADA " & Combo2.Text
Printer.CurrentY = 2.5
Printer.CurrentX = 1
Printer.Print "GRADO";
Printer.CurrentX = 4
Printer.Print "ENERO";
Printer.CurrentX = 6.4
Printer.Print "FEBRERO";
Printer.CurrentX = 8.8
Printer.Print "MARZO";
Printer.CurrentX = 11.2
Printer.Print "ABRIL";
Printer.CurrentX = 13.6
Printer.Print "MAYO";
Printer.CurrentX = 16
Printer.Print "JUNIO"
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "PREJARDIN"
Printer.CurrentX = 1
Printer.Print "JARDIN"
Printer.CurrentX = 1
Printer.Print "TRANSICION"
Printer.CurrentX = 1
Printer.Print "PRIMERO"
Printer.CurrentX = 1
Printer.Print "SEGUNDO"
Printer.CurrentX = 1
Printer.Print "TERCERO"
Printer.CurrentX = 1
Printer.Print "CUARTO"
Printer.CurrentX = 1
Printer.Print "QUINTO"
Printer.CurrentX = 1
Printer.Print "SEXTO"
Printer.CurrentX = 1
Printer.Print "SEPTIMO"
Printer.CurrentX = 1
Printer.Print "OCTAVO"
Printer.CurrentX = 1
Printer.Print "NOVENO"
Printer.CurrentX = 1
Printer.Print "DECIMO"
Printer.CurrentX = 1
Printer.Print "UNDECIMO"
Printer.CurrentX = 1
Printer.Print "DOCE"
Printer.CurrentX = 1
Printer.Print "TRECE"
MATI88.Col = 2
MATI88.Row = 1
Printer.CurrentY = 3
Printer.Print ""
For J = 1 To 16
CX = 4
For k = 1 To 6
Printer.CurrentX = CX
Printer.Print MATI88.Text;
MATI88.Row = MATI88.Row + 1
CX = CX + 2.4
Next k
Printer.Print ""
MATI88.Row = MATI88.Row + 7
Next J
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "TOTALES...";
CX = 4
MATI88.Row = MATI88.Row + 2
For k = 1 To 6
Printer.CurrentX = CX
Printer.Print MATI88.Text;
MATI88.Row = MATI88.Row + 1
CX = CX + 2.4
Next k
Printer.CurrentY = 12.5
Printer.CurrentX = 1
Printer.Print "GRADO";
Printer.CurrentX = 4
Printer.Print "JULIO";
Printer.CurrentX = 6.4
Printer.Print "AGOSTO";
Printer.CurrentX = 8.8
Printer.Print "SEPTIEMBRE";
Printer.CurrentX = 11.8
Printer.Print "OCTUBRE";
Printer.CurrentX = 14.8
Printer.Print "NOVIEMBRE";
Printer.CurrentX = 17.8
Printer.Print "DICIEMBRE"
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "PREJARDIN"
Printer.CurrentX = 1
Printer.Print "JARDIN"
Printer.CurrentX = 1
Printer.Print "TRANSICION"
Printer.CurrentX = 1
Printer.Print "PRIMERO"
Printer.CurrentX = 1
Printer.Print "SEGUNDO"
Printer.CurrentX = 1
Printer.Print "TERCERO"
Printer.CurrentX = 1
Printer.Print "CUARTO"
Printer.CurrentX = 1
Printer.Print "QUINTO"
Printer.CurrentX = 1
Printer.Print "SEXTO"
Printer.CurrentX = 1
Printer.Print "SEPTIMO"
Printer.CurrentX = 1
Printer.Print "OCTAVO"
Printer.CurrentX = 1
Printer.Print "NOVENO"
Printer.CurrentX = 1
Printer.Print "DECIMO"
Printer.CurrentX = 1
Printer.Print "UNDECIMO"
Printer.CurrentX = 1
Printer.Print "DOCE"
Printer.CurrentX = 1
Printer.Print "TRECE"
MATI88.Col = 2
MATI88.Row = 7
Printer.CurrentY = 12.5
Printer.Print ""
Printer.Print ""
For J = 1 To 16
CX = 4
For k = 1 To 6
Printer.CurrentX = CX
Printer.Print MATI88.Text;
MATI88.Row = MATI88.Row + 1
If k > 2 Then
CX = CX + 3
GoTo GISS
End If
CX = CX + 2.4
GISS:
Next k
Printer.Print ""
MATI88.Row = MATI88.Row + 7
Next J
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "TOTALES...";
CX = 4
MATI88.Row = MATI88.Row + 2
For k = 1 To 6
Printer.CurrentX = CX
Printer.Print MATI88.Text;
MATI88.Row = MATI88.Row + 1
If k > 2 Then
CX = CX + 3
GoTo CV
End If
CX = CX + 2.4
CV:
Next k
Printer.EndDoc
End If
End Sub

Private Sub Command6_Click()
'Dim alumno As maestroalum
'Dim pens(1 To 12) As Currency
Label11.Caption = "ESPERE UN MOMENTO POR FAVOR..."
Screen.MousePointer = 11
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
Y = 1
tgp = 0
For k = 1 To 16
    spa = 0
    NAR = FreeFile
    Open Ruta & "pensi.edu" For Random As #NAR Len = 96
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    For J = 1 To 12
        h = 1
        tpm = 0
        While h < I
            Get #NAR, h, alumno
            Get #(NAR - 1), h, pens
            If ((RTrim(alumno.jornada) = Combo3.Text) And (RTrim(alumno.grado) = MATI99.TextMatrix(Y, 1))) And (pens(J) = 0) Then
                tpm = tpm + 1
            End If
            h = h + 1
        Wend
        MATI99.TextMatrix(Y, 2) = tpm
        Y = Y + 1
        spa = spa + tpm
    Next J
    Close #NAR
    Close #(NAR - 1)
    MATI99.TextMatrix(Y, 0) = "SUBTOTAL..."
    MATI99.TextMatrix(Y, 2) = spa
    Y = (k * 13) + 1
    tgp = tgp + spa
Next k
MATI99.TextMatrix(223, 2) = tgp
For k = 1 To 12
    tgp = 0
    Y = k
    For J = 1 To 16
        tgp = tgp + MATI99.TextMatrix(Y, 2)
        Y = (J * 13) + k
    Next J
    Y = Y + 2
    MATI99.TextMatrix(Y, 2) = tgp
Next k
Label11.Caption = ""
Screen.MousePointer = 0
End Sub

Private Sub Command7_Click()
MATI99.Row = 1
MATI99.Col = 2
If MATI99.Text = "" Then
    MsgBox "PRESIONE PRIMERO ACEPTAR", 64, "IMPRIMIR"
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR TOTAL DE PENDIENTES?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR TOTAL PENDIENTES")
If RESP = vbYes Then
Printer.ScaleMode = 7
Printer.CurrentY = 1
Printer.CurrentX = 5
Printer.Print "TOTAL DE PENSIONES PENDIENTES JORNADA " & Combo3.Text
Printer.CurrentY = 2.5
Printer.CurrentX = 1
Printer.Print "GRADO";
Printer.CurrentX = 4
Printer.Print "ENERO";
Printer.CurrentX = 6.4
Printer.Print "FEBRERO";
Printer.CurrentX = 8.8
Printer.Print "MARZO";
Printer.CurrentX = 11.2
Printer.Print "ABRIL";
Printer.CurrentX = 13.6
Printer.Print "MAYO";
Printer.CurrentX = 16
Printer.Print "JUNIO"
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "PREJARDIN"
Printer.CurrentX = 1
Printer.Print "JARDIN"
Printer.CurrentX = 1
Printer.Print "TRANSICION"
Printer.CurrentX = 1
Printer.Print "PRIMERO"
Printer.CurrentX = 1
Printer.Print "SEGUNDO"
Printer.CurrentX = 1
Printer.Print "TERCERO"
Printer.CurrentX = 1
Printer.Print "CUARTO"
Printer.CurrentX = 1
Printer.Print "QUINTO"
Printer.CurrentX = 1
Printer.Print "SEXTO"
Printer.CurrentX = 1
Printer.Print "SEPTIMO"
Printer.CurrentX = 1
Printer.Print "OCTAVO"
Printer.CurrentX = 1
Printer.Print "NOVENO"
Printer.CurrentX = 1
Printer.Print "DECIMO"
Printer.CurrentX = 1
Printer.Print "UNDECIMO"
Printer.CurrentX = 1
Printer.Print "DOCE"
Printer.CurrentX = 1
Printer.Print "TRECE"
MATI99.Col = 2
MATI99.Row = 1
Printer.CurrentY = 3
Printer.Print ""
For J = 1 To 16
CX = 4
For k = 1 To 6
Printer.CurrentX = CX
Printer.Print MATI99.Text;
MATI99.Row = MATI99.Row + 1
CX = CX + 2.4
Next k
Printer.Print ""
MATI99.Row = MATI99.Row + 7
Next J
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "TOTALES...";
CX = 4
MATI99.Row = MATI99.Row + 2
For k = 1 To 6
Printer.CurrentX = CX
Printer.Print MATI99.Text;
MATI99.Row = MATI99.Row + 1
CX = CX + 2.4
Next k
Printer.CurrentY = 12.5
Printer.CurrentX = 1
Printer.Print "GRADO";
Printer.CurrentX = 4
Printer.Print "JULIO";
Printer.CurrentX = 6.4
Printer.Print "AGOSTO";
Printer.CurrentX = 8.8
Printer.Print "SEPTIEMBRE";
Printer.CurrentX = 11.8
Printer.Print "OCTUBRE";
Printer.CurrentX = 14.8
Printer.Print "NOVIEMBRE";
Printer.CurrentX = 17.8
Printer.Print "DICIEMBRE"
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "PREJARDIN"
Printer.CurrentX = 1
Printer.Print "JARDIN"
Printer.CurrentX = 1
Printer.Print "TRANSICION"
Printer.CurrentX = 1
Printer.Print "PRIMERO"
Printer.CurrentX = 1
Printer.Print "SEGUNDO"
Printer.CurrentX = 1
Printer.Print "TERCERO"
Printer.CurrentX = 1
Printer.Print "CUARTO"
Printer.CurrentX = 1
Printer.Print "QUINTO"
Printer.CurrentX = 1
Printer.Print "SEXTO"
Printer.CurrentX = 1
Printer.Print "SEPTIMO"
Printer.CurrentX = 1
Printer.Print "OCTAVO"
Printer.CurrentX = 1
Printer.Print "NOVENO"
Printer.CurrentX = 1
Printer.Print "DECIMO"
Printer.CurrentX = 1
Printer.Print "UNDECIMO"
Printer.CurrentX = 1
Printer.Print "DOCE"
Printer.CurrentX = 1
Printer.Print "TRECE"
MATI99.Col = 2
MATI99.Row = 7
Printer.CurrentY = 12.5
Printer.Print ""
Printer.Print ""
For J = 1 To 16
CX = 4
For k = 1 To 6
Printer.CurrentX = CX
Printer.Print MATI99.Text;
MATI99.Row = MATI99.Row + 1
If k > 2 Then
CX = CX + 3
GoTo GISS2
End If
CX = CX + 2.4
GISS2:
Next k
Printer.Print ""
MATI99.Row = MATI99.Row + 7
Next J
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "TOTALES...";
CX = 4
MATI99.Row = MATI99.Row + 2
For k = 1 To 6
Printer.CurrentX = CX
Printer.Print MATI99.Text;
MATI99.Row = MATI99.Row + 1
If k > 2 Then
CX = CX + 3
GoTo CV2
End If
CX = CX + 2.4
CV2:
Next k
Printer.EndDoc
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Creación y consulta de paz y salvos."
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
   Text7.Text = Format(Text7.Text, "###,###,###.00")
   Call Command1_Click
   Exit Sub
End If
If (KeyAscii = 8) Or (KeyAscii = 46) Then
    If KeyAscii = 46 Then
        If InStr(1, Text7.Text, ".") <> 0 Then
            KeyAscii = 0
        End If
    End If
    Exit Sub
End If
C$ = Chr(KeyAscii)
If C$ < "0" Or C$ > "9" Then
   KeyAscii = 0
   Beep
End If
End Sub

Private Sub TEXT8_KEYPRESS(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command3_Click
End If
If KeyAscii = 8 Then
    Exit Sub
End If
C$ = Chr(KeyAscii)
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Form_Load()
MATI99.Row = 0
MATI99.Col = 0
MATI99.CellForeColor = RGB(255, 255, 255)
MATI99.CellBackColor = RGB(0, 0, 150)
MATI99.ColWidth(0) = 1150
MATI99.Text = "       M E S"
MATI99.Col = 1
MATI99.CellForeColor = RGB(255, 255, 255)
MATI99.CellBackColor = RGB(0, 0, 150)
MATI99.ColWidth(1) = 1100
MATI99.Text = "     GRADO"
MATI99.Col = 2
MATI99.CellForeColor = RGB(255, 255, 255)
MATI99.CellBackColor = RGB(0, 0, 150)
MATI99.ColWidth(2) = 1000
MATI99.Text = "PENDIENTE"
MATI99.Row = 210
MATI99.Col = 0
MATI99.CellForeColor = RGB(255, 255, 255)
MATI99.CellBackColor = RGB(0, 0, 150)
MATI99.Text = "TOTAL X MES"
MATI99.Row = 211
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "ENERO"
MATI99.Row = 212
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "FEBRERO"
MATI99.Row = 213
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "MARZO"
MATI99.Row = 214
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "ABRIL"
MATI99.Row = 215
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "MAYO"
MATI99.Row = 216
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "JUNIO"
MATI99.Row = 217
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "JULIO"
MATI99.Row = 218
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "AGOSTO"
MATI99.Row = 219
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "SEPTIEMBRE"
MATI99.Row = 220
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "OCTUBRE"
MATI99.Row = 221
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "NOVIEMBRE"
MATI99.Row = 222
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "DICIEMBRE"
MATI99.Row = 223
MATI99.CellForeColor = RGB(255, 255, 255)
MATI99.CellBackColor = RGB(0, 0, 150)
MATI99.Text = "T O T A L..."
MATI99.Col = 2
MATI99.CellForeColor = RGB(0, 0, 255)
MATI88.Row = 0
MATI88.Col = 0
MATI88.CellForeColor = RGB(255, 255, 255)
MATI88.CellBackColor = RGB(0, 0, 150)
MATI88.ColWidth(0) = 1150
MATI88.Text = "       M E S"
MATI88.Col = 1
MATI88.CellForeColor = RGB(255, 255, 255)
MATI88.CellBackColor = RGB(0, 0, 150)
MATI88.ColWidth(1) = 1100
MATI88.Text = "     GRADO"
MATI88.Col = 2
MATI88.CellForeColor = RGB(255, 255, 255)
MATI88.CellBackColor = RGB(0, 0, 150)
MATI88.ColWidth(2) = 1350
MATI88.Text = "    VALOR"
MATI88.Row = 210
MATI88.Col = 0
MATI88.CellForeColor = RGB(255, 255, 255)
MATI88.CellBackColor = RGB(0, 0, 150)
MATI88.Text = "TOTAL X MES"
MATI88.Row = 211
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "ENERO"
MATI88.Row = 212
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "FEBRERO"
MATI88.Row = 213
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "MARZO"
MATI88.Row = 214
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "ABRIL"
MATI88.Row = 215
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "MAYO"
MATI88.Row = 216
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "JUNIO"
MATI88.Row = 217
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "JULIO"
MATI88.Row = 218
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "AGOSTO"
MATI88.Row = 219
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "SEPTIEMBRE"
MATI88.Row = 220
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "OCTUBRE"
MATI88.Row = 221
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "NOVIEMBRE"
MATI88.Row = 222
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "DICIEMBRE"
MATI88.Row = 223
MATI88.CellForeColor = RGB(255, 255, 255)
MATI88.CellBackColor = RGB(0, 0, 150)
MATI88.Text = "T O T A L..."
MATI88.Col = 2
MATI88.CellForeColor = RGB(0, 0, 255)
Text9.Text = Format(Date, "mmm/dd/yyyy")
Text7.MaxLength = 14
Text8.MaxLength = 5
MATI77.Row = 0
MATI77.Col = 1
MATI77.ColWidth(1) = 1200
MATI77.Text = "ENERO"
MATI77.Col = 2
MATI77.ColWidth(2) = 1200
MATI77.Text = "FEBRERO"
MATI77.Col = 3
MATI77.ColWidth(3) = 1200
MATI77.Text = "MARZO"
MATI77.Col = 4
MATI77.ColWidth(4) = 1200
MATI77.Text = "ABRIL"
MATI77.Col = 5
MATI77.ColWidth(5) = 1200
MATI77.Text = "MAYO"
MATI77.Col = 6
MATI77.ColWidth(6) = 1200
MATI77.Text = "JUNIO"
MATI77.Col = 7
MATI77.ColWidth(7) = 1200
MATI77.Text = "JULIO"
MATI77.Col = 8
MATI77.ColWidth(8) = 1200
MATI77.Text = "AGOSTO"
MATI77.Col = 9
MATI77.ColWidth(9) = 1200
MATI77.Text = "SEPTIEMBRE"
MATI77.Col = 10
MATI77.ColWidth(10) = 1200
MATI77.Text = "OCTUBRE"
MATI77.Col = 11
MATI77.ColWidth(11) = 1200
MATI77.Text = "NOVIEMBRE"
MATI77.Col = 12
MATI77.ColWidth(12) = 1200
MATI77.Text = "DICIEMBRE"
MATI77.Col = 13
MATI77.ColWidth(13) = 1200
MATI77.Text = "   T O T A L"
MATI77.Row = 1
MATI77.Col = 0
MATI77.Text = "VALOR..."
J = 1
For I = 1 To 16
MATI88.Col = 0
MATI88.Row = J
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "ENERO"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "FEBRERO"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "MARZO"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "ABRIL"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "MAYO"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "JUNIO"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "JULIO"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "AGOSTO"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "SEPTIEMBRE"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "OCTUBRE"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "NOVIEMBRE"
MATI88.Row = MATI88.Row + 1
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Text = "DICIEMBRE"
MATI88.Row = MATI88.Row + 1
MATI88.Text = ""
MATI88.Col = 2
MATI88.CellForeColor = RGB(0, 0, 255)
MATI88.Col = 1
If J = 1 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "PREJARDIN"
Next k
End If
If J = 14 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "JARDIN"
Next k
End If
If J = 27 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "TRANSICION"
Next k
End If
If J = 40 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "PRIMERO"
Next k
End If
If J = 53 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "SEGUNDO"
Next k
End If
If J = 66 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "TERCERO"
Next k
End If
If J = 79 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "CUARTO"
Next k
End If
If J = 92 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "QUINTO"
Next k
End If
If J = 105 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "SEXTO"
Next k
End If
If J = 118 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "SEPTIMO"
Next k
End If
If J = 131 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "OCTAVO"
Next k
End If
If J = 144 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "NOVENO"
Next k
End If
If J = 157 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "DECIMO"
Next k
End If
If J = 170 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "UNDECIMO"
Next k
End If
If J = 183 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "DOCE"
Next k
End If
If J = 196 Then
For k = 1 To 12
MATI88.Row = J + k - 1
MATI88.Text = "TRECE"
Next k
End If
J = J + 13
Next I
J = 1
For I = 1 To 16
MATI99.Col = 0
MATI99.Row = J
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "ENERO"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "FEBRERO"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "MARZO"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "ABRIL"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "MAYO"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "JUNIO"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "JULIO"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "AGOSTO"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "SEPTIEMBRE"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "OCTUBRE"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "NOVIEMBRE"
MATI99.Row = MATI99.Row + 1
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Text = "DICIEMBRE"
MATI99.Row = MATI99.Row + 1
MATI99.Text = ""
MATI99.Col = 2
MATI99.CellForeColor = RGB(0, 0, 255)
MATI99.Col = 1
If J = 1 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "PREJARDIN"
Next k
End If
If J = 14 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "JARDIN"
Next k
End If
If J = 27 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "TRANSICION"
Next k
End If
If J = 40 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "PRIMERO"
Next k
End If
If J = 53 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "SEGUNDO"
Next k
End If
If J = 66 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "TERCERO"
Next k
End If
If J = 79 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "CUARTO"
Next k
End If
If J = 92 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "QUINTO"
Next k
End If
If J = 105 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "SEXTO"
Next k
End If
If J = 118 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "SEPTIMO"
Next k
End If
If J = 131 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "OCTAVO"
Next k
End If
If J = 144 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "NOVENO"
Next k
End If
If J = 157 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "DECIMO"
Next k
End If
If J = 170 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "UNDECIMO"
Next k
End If
If J = 183 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "DOCE"
Next k
End If
If J = 196 Then
For k = 1 To 12
MATI99.Row = J + k - 1
MATI99.Text = "TRECE"
Next k
End If
J = J + 13
Next I
If Dir(Ruta & "pensi.edu") = "" Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
End If
End Sub
