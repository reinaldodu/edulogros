VERSION 5.00
Begin VB.Form AyudaSQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda SQL"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   8295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Copie y pegue la siguiente consulta SQL en el sistema de matrículas para generar el archivo CSV"
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
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "AyudaSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = "SELECT e_nombres, e_apellidos, e_telefono, e_direccion, e_email, e_fnacimiento, e_rh, e_sexo, e_documento, e_eps, grado,p_nombres,p_apellidos, p_telefono,m_nombres, m_apellidos, m_telefono, a_nombres, a_apellidos, a_telefono FROM estudiantes, padres, madres, acudientes, grados WHERE estudiantes.id=p_idest and estudiantes.id=m_idest and estudiantes.id=a_idest and e_grado=grados.id"
End Sub
