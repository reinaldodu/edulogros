VERSION 5.00
Begin VB.Form CORR_ALUM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar datos del estudiante"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   Icon            =   "CORR_ALUM.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text13 
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
      Height          =   285
      Left            =   8640
      TabIndex        =   24
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&GUARDAR"
      Height          =   855
      Left            =   6840
      Picture         =   "CORR_ALUM.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Guarda la información que se muestra en pantalla"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   35
      Top             =   4680
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text10 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "(Digite los últimos cinco (5) números del carnet)."
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3600
         TabIndex        =   37
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "CARNET No."
         Height          =   195
         Left            =   360
         TabIndex        =   36
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "-------------------------------------MODIFICAR DATOS DEL ESTUDIANTE---------------------------"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9255
      Begin VB.TextBox Text22 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6600
         TabIndex        =   23
         Top             =   3600
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "CORR_ALUM.frx":0884
         Left            =   6600
         List            =   "CORR_ALUM.frx":08A0
         TabIndex        =   17
         Text            =   "A +"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3840
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7320
         TabIndex        =   16
         ToolTipText     =   "Año"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6960
         TabIndex        =   15
         ToolTipText     =   "Mes"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8520
         TabIndex        =   18
         Top             =   1080
         Width           =   375
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         TabIndex        =   20
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text8 
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3840
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         TabIndex        =   14
         ToolTipText     =   "Día"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label25 
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
         Left            =   360
         TabIndex        =   51
         Top             =   3720
         Width           =   630
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "TEL. CASA:"
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
         Left            =   360
         TabIndex        =   50
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Label Label23 
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
         Left            =   5280
         TabIndex        =   49
         Top             =   3720
         Width           =   555
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
         Left            =   3360
         TabIndex        =   46
         Top             =   2160
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
         Left            =   360
         TabIndex        =   45
         Top             =   2160
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
         Left            =   3360
         TabIndex        =   44
         Top             =   1680
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
         Left            =   360
         TabIndex        =   43
         Top             =   1680
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "(dd/mm/aaaa)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7920
         TabIndex        =   42
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label15 
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
         Left            =   7800
         TabIndex        =   41
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label13 
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
         Left            =   5280
         TabIndex        =   39
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label12 
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
         Left            =   5280
         TabIndex        =   38
         Top             =   1680
         Width           =   840
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
         Left            =   5280
         TabIndex        =   34
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label8 
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
         Left            =   5280
         TabIndex        =   33
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label7 
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
         Left            =   3360
         TabIndex        =   32
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label6 
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
         Left            =   360
         TabIndex        =   31
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Left            =   360
         TabIndex        =   30
         Top             =   2640
         Width           =   1140
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
         Left            =   5280
         TabIndex        =   29
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label3 
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
         Left            =   5280
         TabIndex        =   28
         Top             =   600
         Width           =   1215
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
         Left            =   360
         TabIndex        =   27
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
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   990
      End
   End
   Begin VB.Label Label22 
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1320
      TabIndex        =   48
      Top             =   240
      Width           =   75
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "CARNET No."
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
      Left            =   120
      TabIndex        =   47
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "MATRICULA No."
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
      Left            =   7080
      TabIndex        =   40
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "CORR_ALUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Combo1_Change()
'If VALI2 = True Then Exit Sub
'If Combo1.Text <> Combo1.List(0) Then
'    Combo1.Text = Combo1.List(0)
'End If
'End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text14.SetFocus
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Luego de haber modificado la información del alumno, de click en Guardar."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text11.SetFocus
End If
If KeyAscii = 8 Then
    Exit Sub
End If
C$ = Format(Chr(KeyAscii), ">")
If C$ <> "M" And C$ <> "F" Then
    KeyAscii = 0
End If
End Sub

Private Sub Text15_Change()
If Len(Text15.Text) = 2 Then
Text16.SetFocus
End If
If Len(Text15.Text) = 0 Then
Text3.SetFocus
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text15.Text) = 1 Then
        Text15.Text = "0" & Text15.Text
    End If
    Text16.SetFocus
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

Private Sub Text16_Change()
If Len(Text16.Text) = 4 Then
Combo1.SetFocus
End If
If Len(Text16.Text) = 0 Then
Text15.SetFocus
End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
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

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text18.SetFocus
End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text19.SetFocus
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

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text20.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text17.SetFocus
End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5.SetFocus
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

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6.SetFocus
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

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_Change()
If Len(Text3.Text) = 2 Then
Text15.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text3.Text) = 1 Then
        Text3.Text = "0" & Text3.Text
    End If
    Text15.SetFocus
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

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text22.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text21.SetFocus
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

Private Sub Text9_KeyPress(KeyAscii As Integer)
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
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
Private Sub TEXT11_KEYPRESS(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text9.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
'If C$ < "0" Or C$ > "9" Then
'    KeyAscii = 0
'    Beep
'End If
End Sub
Private Sub Command1_Click()
'Dim alumno As maestroalum
If Text10.Text = "" Then
    MsgBox "ESCRIBA UN NUMERO DE CARNET", 32, "ADVERTENCIA"
    Text10.SetFocus
    Exit Sub
End If
If Val(Text10.Text) > 32000 Then
    MsgBox "No. DE CARNET INVALIDO", 32, "ADVERTENCIA"
    Text10.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
h = Val(Text10.Text)
If ((h > I - 1) Or (h < 1)) Then
    MsgBox "REGISTRO NO EXISTE", 32
    VERI = 0
    Text10.SetFocus
    Exit Sub
End If
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
    MsgBox "REGISTRO NO EXISTE", 32
    Text10.SetFocus
    Exit Sub
End If
Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
Get #NAR, h, AdiCampo
Close #NAR
Label22.Caption = alumno.n_carnet
Text1.Text = RTrim(alumno.nombres)
Text2.Text = RTrim(alumno.apellidos)
Text11.Text = RTrim(alumno.documento)
Text3.Text = Left(alumno.f_nacimiento, 2)
mm = Right(alumno.f_nacimiento, 7)
Text15.Text = Left(mm, 2)
Text16.Text = Right(alumno.f_nacimiento, 4)
VALI2 = True
Combo1.Text = RTrim(alumno.rh)
VALI2 = False
Text17.Text = RTrim(alumno.padre)
Text18.Text = RTrim(alumno.tel_pa)
Text19.Text = RTrim(alumno.madre)
Text20.Text = RTrim(alumno.tel_ma)
Text5.Text = RTrim(alumno.acudiente)
Text7.Text = RTrim(alumno.tel_acu)
Text21.Text = RTrim(AdiCampo.Tel_casa)
Text6.Text = RTrim(alumno.direccion)
Text8.Text = RTrim(alumno.jornada)
Text9.Text = RTrim(alumno.año_ingre)
Text12.Text = RTrim(alumno.grado)
Text13.Text = RTrim(alumno.n_matricula)
Text14.Text = RTrim(alumno.sexo)
Text4.Text = RTrim(AdiCampo.salud)
Text22.Text = RTrim(AdiCampo.email)
Text1.SetFocus
VERI = 1
End Sub
Private Sub Command2_Click()
'Dim alumno As maestroalum
If VERI = 0 Then
    MsgBox "SELECCIONE PRIMERO EL REGISTRO A CORREGIR", vbCritical, "ADVERTENCIA"
    Text10.SetFocus
    Exit Sub
End If
If (RTrim(Text1.Text) = "") Or (RTrim(Text2.Text) = "") Or (RTrim(Text6.Text) = "") Or (RTrim(Text9.Text) = "") Or (RTrim(Text3.Text) = "") Or (RTrim(Text15.Text) = "") Or (RTrim(Text16.Text) = "") Or (RTrim(Text13.Text) = "") Then
    MsgBox "INFORMACION INCOMPLETA PARA GUARDAR", 16, "GUARDAR"
    Exit Sub
End If
If (Val(Text3.Text) < 1) Or (Val(Text3.Text) > 31) Then
    MsgBox "DIA INVALIDO", 48, "FECHA DE NACIMIENTO"
    Text3.SetFocus
    Exit Sub
End If
If (Val(Text15.Text) < 1) Or (Val(Text15.Text) > 12) Then
    MsgBox "MES INVALIDO", 48, "FECHA DE NACIMIENTO"
    Text15.SetFocus
    Exit Sub
End If
If Val(Text16.Text) < 1900 Then
    MsgBox "AÑO INVALIDO", 48, "FECHA DE NACIMIENTO"
    Text16.SetFocus
    Exit Sub
End If
If Len(Text3.Text) = 1 Then
    Text3.Text = "0" & Text3.Text
End If
If Len(Text15.Text) = 1 Then
    Text15.Text = "0" & Text15.Text
End If
Text14.Text = Format(Text14.Text, ">")
If Text14.Text <> "M" And Text14.Text <> "F" Then
    MsgBox "SEXO DEBE SER (M) ó (F)", 16, "GUARDAR"
    Text14.SetFocus
    Exit Sub
End If
If Val(Text9.Text) < 1900 Then
    MsgBox "AÑO DE INGRESO INVALIDO", 48, "ADVERTENCIA"
    Text9.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA GUARDAR LA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR REGISTRO")
If RESP = vbYes Then
    alumno.n_carnet = Label22.Caption
    alumno.nombres = Format(Text1.Text, ">")
    alumno.apellidos = Format(Text2.Text, ">")
    alumno.documento = Text11.Text
    alumno.f_nacimiento = Text3.Text & "/" & Text15.Text & "/" & Text16.Text
    alumno.rh = Combo1.Text
    alumno.acudiente = Format(Text5.Text, ">")
    alumno.tel_acu = Text7.Text
    alumno.padre = Format(Text17.Text, ">")
    alumno.tel_pa = Text18.Text
    alumno.madre = Format(Text19.Text, ">")
    alumno.tel_ma = Text20.Text
    alumno.direccion = Text6.Text
    alumno.jornada = Text8.Text
    alumno.año_ingre = Text9.Text
    alumno.grado = Text12.Text
    alumno.n_matricula = Val(Text13.Text)
    alumno.sexo = Text14.Text
    AdiCampo.salud = Format(Text4.Text, ">")
    AdiCampo.Tel_casa = Text21.Text
    AdiCampo.email = Text22.Text
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Put #NAR, h, alumno
    Close #NAR
    Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
    Put #NAR, h, AdiCampo
    Close #NAR
End If
End Sub

Private Sub TEXT13_KEYPRESS(KeyAscii As Integer)
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
VERI = 0
Text1.MaxLength = 20
Text2.MaxLength = 20
Text11.MaxLength = 12
Text3.MaxLength = 2
Text4.MaxLength = 20
Text5.MaxLength = 30
Text6.MaxLength = 40
Text7.MaxLength = 12
Text9.MaxLength = 4
Text10.MaxLength = 5
Text13.MaxLength = 5
Text14.MaxLength = 1
Text15.MaxLength = 2
Text16.MaxLength = 4
Text17.MaxLength = 30
Text18.MaxLength = 12
Text19.MaxLength = 30
Text20.MaxLength = 12
Text21.MaxLength = 12
End Sub
