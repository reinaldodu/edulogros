VERSION 5.00
Begin VB.Form RETI_PRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base de datos de profesores retirados"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7230
   Icon            =   "RETI_PRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   5160
      Picture         =   "RETI_PRO.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Imprime los datos personales mostrados en pantalla"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text12 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buscar por:"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   4335
      Begin VB.Frame Frame4 
         Caption         =   "Año de retiro"
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   2520
         TabIndex        =   30
         Top             =   240
         Width           =   1695
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Left            =   600
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Ok"
            Height          =   375
            Left            =   480
            TabIndex        =   25
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Año:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   330
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   600
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text11 
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
         Left            =   600
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cédula"
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2175
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "No."
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   255
         End
      End
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
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.TextBox Text10 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label10 
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
         Left            =   3720
         TabIndex        =   10
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label9 
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
         Left            =   3720
         TabIndex        =   9
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Label8 
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
         Left            =   3720
         TabIndex        =   8
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "TITULO:"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   1080
         Width           =   750
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
         Left            =   3720
         TabIndex        =   6
         Top             =   480
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
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CEDULA No."
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
         TabIndex        =   3
         Top             =   1680
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
         TabIndex        =   2
         Top             =   1080
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
         TabIndex        =   1
         Top             =   480
         Width           =   990
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Total retirados..."
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5160
      TabIndex        =   32
      Top             =   3840
      Width           =   1140
   End
End
Attribute VB_Name = "RETI_PRO"
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
Call Command3_Click
End If
End Sub

Private Sub Command1_Click()
'Dim proti As pro_reti
If Text11.Text = "" Then
    MsgBox "ESCRIBA EL No. DE CEDULA", 32, "BUSCAR"
    Text11.SetFocus
    Exit Sub
End If
CED = Text11.Text
QQ = 0
NAR = FreeFile
Open Ruta & "retipro.edu" For Random As #NAR Len = Len(proti)
While Not EOF(NAR)
QQ = QQ + 1
Get #NAR, QQ, proti
If RTrim(CED) = RTrim(proti.documento) Then
Text1.Text = proti.nombres
Text2.Text = proti.apellidos
Text3.Text = proti.documento
Text4.Text = proti.rh
Text5.Text = proti.direccion
Text6.Text = proti.Telefono
Text7.Text = proti.especiali
Text8.Text = proti.escalafon
Text9.Text = proti.año_ingre
Text10.Text = proti.año_retir
Close #NAR
Exit Sub
End If
Wend
Close #NAR
MsgBox "REGISTRO NO ENCONTRADO", 48, "BUSCAR"
Text11.SetFocus
End Sub

Private Sub Command2_Click()
'Dim ini As inicio
If Text1.Text = "" Then
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
'Printer.Font.Underline = True
Printer.CurrentY = 1
Printer.CurrentX = 8.5
Printer.Font.Size = 12
Printer.Print "PROFESOR RETIRADO"
Printer.Print ""
Printer.Print ""
Printer.Font.Size = 10
Printer.CurrentX = 1
Printer.Print ini.nombre
Printer.CurrentY = 4
Printer.CurrentX = 1
Printer.Print "NOMBRES: " & Text1.Text
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "APELLIDOS: " & Text2.Text
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "DOCUMENTO: " & Text3.Text
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "FACTOR R-H: " & Text4.Text
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "TELEFONO: " & Text6.Text
Printer.Print ""
Printer.CurrentY = 4
Printer.CurrentX = 8.5
Printer.Print "DIRECCION: " & Text5.Text
Printer.Print ""
Printer.CurrentX = 8.5
Printer.Print "TITULO: " & Text7.Text
Printer.Print ""
Printer.CurrentX = 8.5
Printer.Print "ESCALAFON: " & Text8.Text
Printer.Print ""
Printer.CurrentX = 8.5
Printer.Print "AÑO DE INGRESO: " & Text9.Text
Printer.Print ""
Printer.CurrentX = 8.5
Printer.Print "AÑO DE RETIRO: " & Text10.Text
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Command3_Click()
'Dim proti As pro_reti
p = 1
QQ = 0
NAR = FreeFile
Open Ruta & "retipro.edu" For Random As #NAR Len = Len(proti)
While Not EOF(NAR)
    QQ = QQ + 1
    Get #NAR, QQ, proti
    If Combo1.Text = RTrim(proti.año_retir) Then
        BUSQ_RETIPRO.MATI30.Rows = p + 1
        BUSQ_RETIPRO.MATI30.TextMatrix(p, 0) = RTrim(proti.nombres) & " " & RTrim(proti.apellidos)
        BUSQ_RETIPRO.MATI30.TextMatrix(p, 1) = RTrim(proti.documento)
        BUSQ_RETIPRO.MATI30.TextMatrix(p, 2) = RTrim(proti.direccion)
        BUSQ_RETIPRO.MATI30.TextMatrix(p, 3) = RTrim(proti.Telefono)
        p = p + 1
    End If
Wend
Close #NAR
If p = 1 Then
    MsgBox "NO SE ENCONTRARON REGISTROS", 16, "BUSCAR POR"
    Exit Sub
End If
BUSQ_RETIPRO.Frame1.Caption = "AÑO DE RETIRO: " & Combo1.Text
BUSQ_RETIPRO.Label1.Caption = "REGISTROS ENCONTRADOS = " & (p - 1)
BUSQ_RETIPRO.Show
End Sub

Private Sub Command4_Click()
BUSQ_RETIPRO.Show
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Base de datos de profesores retirados."
End Sub

Private Sub TEXT11_KEYPRESS(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Form_Load()
For I = 1998 To 2100
Combo1.AddItem I
Next I
Combo1.Text = Combo1.List(0)
NAR = FreeFile
Open Ruta & "conrepro.edu" For Input As #NAR
Input #NAR, zu
Close #NAR
Text12.Text = zu - 1
Text11.MaxLength = 10
End Sub
