VERSION 5.00
Begin VB.Form CONTRASEÑA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso - Edulogros"
   ClientHeight    =   4215
   ClientLeft      =   2550
   ClientTop       =   1890
   ClientWidth     =   4335
   Icon            =   "CONTRASEÑA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cambiar contraseña"
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   4095
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
         Left            =   2280
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text4 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   1560
         Width           =   975
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
         ForeColor       =   &H0000FFFF&
         Height          =   320
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Escriba su nuevo usuario:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1830
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Confirme su nueva clave:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Escriba su nueva clave:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1710
      End
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
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
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
      ForeColor       =   &H0000FFFF&
      Height          =   320
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contraseña de acceso"
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4095
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Escriba su contraseña:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Escriba su usuario:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "CONTRASEÑA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim contra As clave
Dim NOM_CLAV As String, PASS_CLAV As String, NEW_NOM As String, NEW_PASS As String

If Text1.Text = "" Then
    MsgBox "ESCRIBA EL NOMBRE DE USUARIO", 16, "CONTRASEÑA"
    Text1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "ESCRIBA SU PASSWORD", 16, "CONTRASEÑA"
    Text2.SetFocus
    Exit Sub
End If
CLA = 0
NAR = FreeFile
Open App.Path & "\clase.edu" For Random As #NAR Len = Len(contra)
While Not EOF(NAR)
CLA = CLA + 1
Get #NAR, CLA, contra
NOM_CLAV = ""
PASS_CLAV = ""

For I = 1 To Len(Trim(contra.nombre)) Step 3
    NOM_CLAV = NOM_CLAV & Chr(Val(Mid(contra.nombre, I, 3)))
Next
For I = 1 To Len(Trim(contra.PASSW)) Step 3
    PASS_CLAV = PASS_CLAV & Chr(Val(Mid(contra.PASSW, I, 3)))
Next

If (UCase(Text1) = NOM_CLAV) And (UCase(Text2) = PASS_CLAV) Then
    Close #NAR
    clacerr = True
    Unload Me
    ENTRADA.Show
    Exit Sub
End If
Wend
Close #NAR
If CLA <= 3 Then
RESP = MsgBox("DESEA CREAR ESTA CONTRASEÑA?", vbYesNo + vbQuestion + vbDefaultButton2, "CONTRASEÑA")
If RESP = vbYes Then
    Open App.Path & "\clase.edu" For Random As #NAR Len = Len(contra)
    'contra.nombre = RTrim(Text1.Text)
    'contra.PASSW = RTrim(Text2.Text)
    NOM_CLAV = UCase(Trim(Text1))
    PASS_CLAV = UCase(Trim(Text2))
    
    NEW_NOM = ""
    NEW_PASS = ""
    
    For I = 1 To Len(NOM_CLAV)
        NEW_NOM = NEW_NOM & "0" & Asc(Mid(NOM_CLAV, I, 1))
    Next
    
    For I = 1 To Len(PASS_CLAV)
        NEW_PASS = NEW_PASS & "0" & Asc(Mid(PASS_CLAV, I, 1))
    Next
    contra.nombre = NEW_NOM
    contra.PASSW = NEW_PASS
    Put #NAR, CLA, contra
    Close #NAR
    End If
Else
    malo = malo + 1
    MsgBox "CONTRASEÑA INCORRECTA", 16, "BIENVENIDOS A EDULOGROS"
    Text2.SetFocus
    'SendKeys "{Home}+{End}"
    If malo = 3 Then
        End
    End If
End If
End Sub

Private Sub Command2_Click()
'Dim contra As clave
Dim NOM_CLAV As String, PASS_CLAV As String, NEW_NOM As String, NEW_PASS As String

If Text1.Text = "" Then
    MsgBox "ESCRIBA EL NOMBRE DE USUARIO PARA CAMBIAR DE CONTRASEÑA", 16, "CONTRASEÑA"
    Text1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "ESCRIBA SU PASSWORD PARA CAMBIAR DE CONTRASEÑA", 16, "CONTRASEÑA"
    Text2.SetFocus
    Exit Sub
End If
If Text3.Text = "" Then
    MsgBox "ESCRIBA SU NUEVO PASSWORD", 16, "CONTRASEÑA"
    Text3.SetFocus
    Exit Sub
End If
If Text4.Text = "" Then
    MsgBox "CONFIRME SU NUEVO PASSWORD", 16, "CONTRASEÑA"
    Text4.SetFocus
    Exit Sub
End If
If Text5.Text = "" Then
    MsgBox "ESCRIBA SU NUEVO NOMBRE", 16, "CONTRASEÑA"
    Text5.SetFocus
    Exit Sub
End If
If Text3.Text <> Text4.Text Then
    MsgBox "CONFIRMACION DE CONTRASEÑA NO COINCIDE", 16, "CONTRASEÑA"
    Text4.Text = ""
    Text4.SetFocus
    Exit Sub
End If
RESP = MsgBox("REALMENTE DESEA CAMBIAR LA CONTRASEÑA?", vbYesNo + vbQuestion + vbDefaultButton2, "BIENVENIDOS A EDULOGROS")
If RESP = vbYes Then
CLA = 0
NAR = FreeFile
Open App.Path & "\clase.edu" For Random As #NAR Len = Len(contra)
While Not EOF(NAR)
CLA = CLA + 1
Get #NAR, CLA, contra

NOM_CLAV = ""
PASS_CLAV = ""

For I = 1 To Len(Trim(contra.nombre)) Step 3
    NOM_CLAV = NOM_CLAV & Chr(Val(Mid(contra.nombre, I, 3)))
Next
For I = 1 To Len(Trim(contra.PASSW)) Step 3
    PASS_CLAV = PASS_CLAV & Chr(Val(Mid(contra.PASSW, I, 3)))
Next

If (UCase(Text1) = NOM_CLAV) And (UCase(Text2) = PASS_CLAV) Then
    NOM_CLAV = UCase(Trim(Text5))
    PASS_CLAV = UCase(Trim(Text3))
    
    NEW_NOM = ""
    NEW_PASS = ""
    
    For I = 1 To Len(NOM_CLAV)
        NEW_NOM = NEW_NOM & "0" & Asc(Mid(NOM_CLAV, I, 1))
    Next
    
    For I = 1 To Len(PASS_CLAV)
        NEW_PASS = NEW_PASS & "0" & Asc(Mid(PASS_CLAV, I, 1))
    Next
    contra.nombre = NEW_NOM
    contra.PASSW = NEW_PASS
    Put #NAR, CLA, contra
    Close #NAR
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text1.SetFocus
    Exit Sub
End If


Wend
Close #NAR
malo = malo + 1
MsgBox "CONTRASEÑA INCORRECTA", 16, "CONTRASEÑA"
Text2.Text = ""
Text2.SetFocus
If malo = 3 Then
Unload Me
End If
End If
End Sub

Private Sub Form_Load()
Text1.MaxLength = 15
Text2.MaxLength = 15
Text3.MaxLength = 15
Text4.MaxLength = 15
Text5.MaxLength = 15
malo = 0
clacerr = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If clacerr = False Then
    End
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command2_Click
End If
End Sub
