VERSION 5.00
Begin VB.Form CONFIG_DESEMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración desempeños por grado"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Abreviaturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   3855
      Begin VB.TextBox txt4_recup 
         Height          =   285
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   35
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txt3_recup 
         Height          =   285
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   34
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt2_recup 
         Height          =   285
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   33
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txt1_recup 
         Height          =   285
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txt4_desemp 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   15
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txt3_desemp 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt2_desemp 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txt1_desemp 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Recup."
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
         TabIndex        =   31
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desempeño Bajo:        --->"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desempeño Básico:    --->"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Desempeño Alto:         --->"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Desempeño Superior:  --->"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1845
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "CONFIG_DESEMP.frx":0000
      Left            =   2280
      List            =   "CONFIG_DESEMP.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rangos porcentuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      Begin VB.TextBox txt3_porcent 
         Height          =   285
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txt2_porcent 
         Height          =   285
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   18
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt1_porcent 
         Height          =   285
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "%Máx."
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
         TabIndex        =   30
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "%Min."
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
         Left            =   2040
         TabIndex        =   29
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3360
         TabIndex        =   28
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3360
         TabIndex        =   27
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3360
         TabIndex        =   26
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   2760
         TabIndex        =   25
         Top             =   600
         Width           =   390
      End
      Begin VB.Label lbl4_porcent 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   2040
         Width           =   210
      End
      Begin VB.Label lbl3_porcent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2040
         TabIndex        =   22
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label lbl2_porcent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2040
         TabIndex        =   21
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label lbl1_porcent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2040
         TabIndex        =   20
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desempeño Bajo:        --->"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desempeño Básico:    --->"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1830
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desempeño Alto:         --->"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desempeño Superior:  --->"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1845
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Grado:"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "CONFIG_DESEMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
NAR = FreeFile
Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
Get #NAR, Combo1.ListIndex + 1, confdesemp
Close #NAR
txt1_porcent = confdesemp.rango(1)
txt2_porcent = confdesemp.rango(2)
txt3_porcent = confdesemp.rango(3)
txt1_desemp = Trim(confdesemp.desemp(1))
txt2_desemp = Trim(confdesemp.desemp(2))
txt3_desemp = Trim(confdesemp.desemp(3))
txt4_desemp = Trim(confdesemp.desemp(4))
txt1_recup = Trim(confdesemp.recupera(1))
txt2_recup = Trim(confdesemp.recupera(2))
txt3_recup = Trim(confdesemp.recupera(3))
txt4_recup = Trim(confdesemp.recupera(4))
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If (Val(txt1_porcent) <= Val(txt3_porcent)) Or (Val(txt1_porcent) <= Val(txt2_porcent)) Then
    MsgBox "El porcentaje para el desempeño alto no puede ser menor o igual al del desempeño bajo o básico", 16, "Desempeños"
    Exit Sub
End If
If (Val(txt2_porcent) <= Val(txt3_porcent)) Then
    MsgBox "El porcentaje para el desempeño básico no puede ser menor o igual al del desempeño bajo", 16, "Desempeños"
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
confdesemp.rango(1) = txt1_porcent
confdesemp.rango(2) = txt2_porcent
confdesemp.rango(3) = txt3_porcent
confdesemp.desemp(1) = Trim(txt1_desemp)
confdesemp.desemp(2) = Trim(txt2_desemp)
confdesemp.desemp(3) = Trim(txt3_desemp)
confdesemp.desemp(4) = Trim(txt4_desemp)
confdesemp.recupera(1) = Trim(txt1_recup)
confdesemp.recupera(2) = Trim(txt2_recup)
confdesemp.recupera(3) = Trim(txt3_recup)
confdesemp.recupera(4) = Trim(txt4_recup)
confdesemp.grado = Trim(Combo1.Text)
Put #NAR, Combo1.ListIndex + 1, confdesemp
Close #NAR
MsgBox "Configuración guardada", 64, "Desempeños"
End Sub

Private Sub Form_Load()
Combo1 = Combo1.List(0)
End Sub

Private Sub txt1_porcent_Change()
lbl1_porcent = Val(txt1_porcent) + 1 & "%"
End Sub

Private Sub txt1_porcent_KeyPress(KeyAscii As Integer)
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txt2_porcent_Change()
lbl2_porcent = Val(txt2_porcent) + 1 & "%"
End Sub

Private Sub txt2_porcent_KeyPress(KeyAscii As Integer)
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txt3_porcent_Change()
lbl3_porcent = Val(txt3_porcent) + 1 & "%"
End Sub

Private Sub txt3_porcent_KeyPress(KeyAscii As Integer)
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub
