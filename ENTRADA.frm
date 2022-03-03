VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ENTRADA 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6795
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9510
   Icon            =   "ENTRADA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "datalu"
            Object.ToolTipText     =   "Adicionar y consultar estudiantes"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "datapro"
            Object.ToolTipText     =   "Adicionar y consultar profesores"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "matric"
            Object.ToolTipText     =   "Matrículas"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "alumexi"
            Object.ToolTipText     =   "Estudiantes existentes por grado"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "directele"
            Object.ToolTipText     =   "Directorio telefónico por grupo"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "congru"
            Object.ToolTipText     =   "Consultar grupo"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "conare"
            Object.ToolTipText     =   "Consultar materias"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "aregra"
            Object.ToolTipText     =   "Materias por grupo"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "adiobs"
            Object.ToolTipText     =   "Adicionar y consultar logros"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "adiobal"
            Object.ToolTipText     =   "Grabar desempeños"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "agreobs"
            Object.ToolTipText     =   "Adicionar observaciones"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "conobal"
            Object.ToolTipText     =   "Consultar e imprimir boletines"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "impobal"
            Description     =   ""
            Object.ToolTipText     =   "Crear cuentas de profesores de la red LAN"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "esta"
            Object.ToolTipText     =   "Control de logros perdidos"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "contotal"
            Object.ToolTipText     =   "Control de totales"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "disdapro"
            Object.ToolTipText     =   "Crear datos profesor"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6420
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14384
            MinWidth        =   3457
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   600
      Picture         =   "ENTRADA.frx":27A2
      ScaleHeight     =   5295
      ScaleWidth      =   8325
      TabIndex        =   1
      Top             =   960
      Width           =   8325
      Begin ComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   4200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   21
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":8F408
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":8F722
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":8FA3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":8FD56
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":90070
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":9038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":906A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":90FBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":912D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":915F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":9190C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":91C26
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":92540
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":92DA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":93594
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":93EAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":946E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":949FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":94D14
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":9502E
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ENTRADA.frx":95348
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu dat 
      Caption         =   "&Datos"
      Begin VB.Menu inici 
         Caption         =   "I&niciales"
      End
      Begin VB.Menu alu 
         Caption         =   "&Estudiantes"
         Begin VB.Menu adi 
            Caption         =   "&Nuevo"
            Shortcut        =   ^A
         End
         Begin VB.Menu cons 
            Caption         =   "&Consultar"
            Shortcut        =   ^C
         End
         Begin VB.Menu Corr 
            Caption         =   "&Modificar"
         End
         Begin VB.Menu retalma 
            Caption         =   "&Borrar"
         End
      End
      Begin VB.Menu prof 
         Caption         =   "&Profesores"
         Begin VB.Menu adipro 
            Caption         =   "&Nuevo"
            Shortcut        =   ^D
         End
         Begin VB.Menu conpro 
            Caption         =   "&Consultar"
            Shortcut        =   ^T
         End
         Begin VB.Menu correpro 
            Caption         =   "&Modificar"
         End
         Begin VB.Menu retipro 
            Caption         =   "&Borrar"
         End
      End
      Begin VB.Menu raya 
         Caption         =   "-"
      End
      Begin VB.Menu configfolio 
         Caption         =   "Configurar número de folio"
      End
      Begin VB.Menu raya2 
         Caption         =   "-"
      End
      Begin VB.Menu salis 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu alilis 
      Caption         =   "&Estudiantes"
      Begin VB.Menu infoadi 
         Caption         =   "Información adicional"
      End
      Begin VB.Menu rayainfoadi 
         Caption         =   "-"
      End
      Begin VB.Menu alumn 
         Caption         =   "&Matrículas"
         Shortcut        =   ^M
      End
      Begin VB.Menu linear 
         Caption         =   "-"
      End
      Begin VB.Menu gralu 
         Caption         =   "&Existentes por grado"
      End
      Begin VB.Menu terray 
         Caption         =   "-"
      End
      Begin VB.Menu singrupen 
         Caption         =   "&Pendientes y Sin Grupo"
      End
      Begin VB.Menu lincva 
         Caption         =   "-"
      End
      Begin VB.Menu cvnec 
         Caption         =   "C&onsultas opcionales"
         Begin VB.Menu cvnoap 
            Caption         =   "&Nombres y Apellidos"
         End
         Begin VB.Menu cvraca 
            Caption         =   "&Rango de carnets"
         End
         Begin VB.Menu cveda 
            Caption         =   "&Edad"
         End
         Begin VB.Menu cvrh 
            Caption         =   "R.&H."
         End
         Begin VB.Menu cvsx 
            Caption         =   "&Sexo"
         End
         Begin VB.Menu cvaing 
            Caption         =   "&Año de ingreso"
         End
      End
   End
   Begin VB.Menu grup 
      Caption         =   "&Grupo"
      Begin VB.Menu cregru 
         Caption         =   "Cr&ear"
         Shortcut        =   ^R
      End
      Begin VB.Menu congru 
         Caption         =   "&Consultar"
         Shortcut        =   ^U
      End
      Begin VB.Menu grumodi 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu AliasGroups 
         Caption         =   "Alias"
      End
      Begin VB.Menu impr 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu cuarray 
         Caption         =   "-"
      End
      Begin VB.Menu direc 
         Caption         =   "&Grupos por grado"
      End
      Begin VB.Menu direforaya 
         Caption         =   "-"
      End
      Begin VB.Menu directelefo 
         Caption         =   "&Directorio telefónico"
      End
   End
   Begin VB.Menu area 
      Caption         =   "&Materias"
      Begin VB.Menu creare 
         Caption         =   "Cr&ear"
      End
      Begin VB.Menu conare 
         Caption         =   "&Consultar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu corrare 
         Caption         =   "C&orregir"
      End
      Begin VB.Menu casra 
         Caption         =   "-"
      End
      Begin VB.Menu arepogr 
         Caption         =   "Materias por grupo"
         Shortcut        =   ^F
      End
      Begin VB.Menu Orargru 
         Caption         =   "Ordenar materias por grupo"
      End
   End
   Begin VB.Menu aluob 
      Caption         =   "&Boletín"
      Begin VB.Menu obse 
         Caption         =   "&Logros e indicadores"
      End
      Begin VB.Menu graba_PtjLgr 
         Caption         =   "Grabar porcentajes de logros"
      End
      Begin VB.Menu grabadesemp 
         Caption         =   "Grabar desempeños"
      End
      Begin VB.Menu adiciobal 
         Caption         =   "&Grabar observaciones"
         Shortcut        =   ^G
      End
      Begin VB.Menu rayacons2 
         Caption         =   "-"
      End
      Begin VB.Menu imalob 
         Caption         =   "Impr&esión y consulta de boletines"
         Shortcut        =   ^P
      End
      Begin VB.Menu linea3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu refinal 
         Caption         =   "&Informe final"
         Visible         =   0   'False
      End
      Begin VB.Menu linea4 
         Caption         =   "-"
      End
      Begin VB.Menu infofinal 
         Caption         =   "Comentarios en el &Boletín"
      End
      Begin VB.Menu commdesempp 
         Caption         =   "Comentarios por desempeño"
      End
      Begin VB.Menu infrfin 
         Caption         =   "Co&mentarios en el informe final"
      End
   End
   Begin VB.Menu contt 
      Caption         =   "&Control"
      Begin VB.Menu cambibase 
         Caption         =   "Cambiar Base de datos"
      End
      Begin VB.Menu raypapro 
         Caption         =   "-"
      End
      Begin VB.Menu porlogros44 
         Caption         =   "Porcentaje de logros"
      End
      Begin VB.Menu paradesemp 
         Caption         =   "Parámetros de desempeños"
      End
      Begin VB.Menu menparapro 
         Caption         =   "P&arámetros de promoción"
      End
      Begin VB.Menu bloquedesemp 
         Caption         =   "Bloquear periodos"
      End
      Begin VB.Menu rayacambase 
         Caption         =   "-"
      End
      Begin VB.Menu crearuser 
         Caption         =   "Cuentas de usuario en red"
      End
      Begin VB.Menu rayausers 
         Caption         =   "-"
      End
      Begin VB.Menu coprofeco 
         Caption         =   "&Datos sistema profesor"
         Begin VB.Menu disdaprofe 
            Caption         =   "&Crear datos profesor"
         End
         Begin VB.Menu badispro 
            Caption         =   "&Bajar datos profesor"
         End
         Begin VB.Menu controlcerradas 
            Caption         =   "Control de planillas"
         End
         Begin VB.Menu regdataprof 
            Caption         =   "Registro datos profesores"
         End
      End
      Begin VB.Menu prolprols 
         Caption         =   "-"
      End
      Begin VB.Menu Enlace_planea 
         Caption         =   "Planeación"
         Begin VB.Menu eje_tem_cont 
            Caption         =   "Ejes temáticos y contenidos"
         End
         Begin VB.Menu Enlace_comp 
            Caption         =   "Competencias"
         End
         Begin VB.Menu planea_sema 
            Caption         =   "Planeación semanal"
         End
      End
      Begin VB.Menu Enlace_Proyectos 
         Caption         =   "Proyectos"
         Visible         =   0   'False
      End
      Begin VB.Menu mitadtri 
         Caption         =   "Mitad de periodo"
         Visible         =   0   'False
         Begin VB.Menu mit_impr 
            Caption         =   "Impresión reportes"
         End
         Begin VB.Menu txt_encab2 
            Caption         =   "Texto encabezado del boletín"
         End
         Begin VB.Menu separa_bol2 
            Caption         =   "-"
         End
         Begin VB.Menu mit_logreap 
            Caption         =   "Logros perdidos y reaprendizaje"
         End
      End
      Begin VB.Menu linea_mitrimes 
         Caption         =   "-"
      End
      Begin VB.Menu logropen 
         Caption         =   "&Informes generales"
         Begin VB.Menu infoareas 
            Caption         =   "Porcentaje logros y desempeños"
         End
         Begin VB.Menu logporperiod 
            Caption         =   "Logros perdidos y reaprendizaje"
         End
      End
      Begin VB.Menu reinal 
         Caption         =   "-"
      End
      Begin VB.Menu infespe 
         Caption         =   "Informes específicos"
         Begin VB.Menu infoareas2 
            Caption         =   "Porcentajes logros y desempeños"
         End
         Begin VB.Menu logporperiod2 
            Caption         =   "Logros perdidos y reaprendizaje"
         End
      End
      Begin VB.Menu rayainf2 
         Caption         =   "-"
      End
      Begin VB.Menu CarnWord 
         Caption         =   "Carnets en &Word"
      End
      Begin VB.Menu Licars 
         Caption         =   "-"
      End
      Begin VB.Menu pensi 
         Caption         =   "P&ensiones"
         Shortcut        =   {F3}
      End
      Begin VB.Menu gtsf 
         Caption         =   "-"
      End
      Begin VB.Menu tota 
         Caption         =   "&Totales"
         Shortcut        =   {F4}
      End
      Begin VB.Menu saby 
         Caption         =   "-"
      End
      Begin VB.Menu histo 
         Caption         =   "Creación del &historial"
      End
      Begin VB.Menu nvlin 
         Caption         =   "-"
      End
      Begin VB.Menu inian 
         Caption         =   "&Inicio de año"
      End
   End
   Begin VB.Menu ayud 
      Caption         =   "Ay&uda"
      Begin VB.Menu ayuedul 
         Caption         =   "Ayuda de &Edulogros"
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu ayurap 
         Caption         =   "Ayuda rápida de &menús"
         Shortcut        =   {F1}
      End
      Begin VB.Menu liayuweb 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu ayulinea 
         Caption         =   "Edulogros en &Internet"
         Visible         =   0   'False
      End
      Begin VB.Menu rayita 
         Caption         =   "-"
      End
      Begin VB.Menu acerca 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu vari 
      Caption         =   "varios"
      Visible         =   0   'False
      Begin VB.Menu ggg 
         Caption         =   "I&mprimir materias"
      End
      Begin VB.Menu bb 
         Caption         =   "Imprimir gr&upo"
      End
      Begin VB.Menu aa 
         Caption         =   "Imprimir &boletines por grupo"
      End
      Begin VB.Menu mnspimp 
         Caption         =   "-"
      End
      Begin VB.Menu cc 
         Caption         =   "Grabar logros"
      End
      Begin VB.Menu grabdesemp2 
         Caption         =   "Grabar desempeños"
      End
      Begin VB.Menu grabogru 
         Caption         =   "Gr&abar observaciones"
      End
      Begin VB.Menu mnspgrab 
         Caption         =   "-"
      End
      Begin VB.Menu ccc 
         Caption         =   "&Consultar boletín"
      End
      Begin VB.Menu izconsugrup 
         Caption         =   "Consultar &grupo"
      End
      Begin VB.Menu cocoalu 
         Caption         =   "Con&sultar estudiante"
      End
      Begin VB.Menu cocoproro 
         Caption         =   "Co&nsultar profesor"
      End
   End
End
Attribute VB_Name = "ENTRADA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aa_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CONS_NOTA.Show
End If
End Sub

Private Sub ACERCA_Click()
ACERCADE.Show 1
End Sub

Private Sub adi_Click()
I = 0
PASSW.Show 1
If I = 1 Then
BASE_ALUM.Show
BASE_ALUM.Text2.SetFocus
End If
End Sub

Private Sub adiciobal_Click()
I = 0
PASSW.Show 1
If I = 1 Then
GRABAR_OBSER.Show
End If
End Sub

Private Sub adipro_Click()
I = 0
PASSW.Show 1
If I = 1 Then
BASE_PROF.Show
End If
End Sub

Private Sub AliasGroups_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    Alias_Grupo.Show 1
End If
End Sub

Private Sub alumn_Click()
I = 0
PASSW.Show 1
If I = 1 Then
MATRICULA.Show 1
End If
End Sub

Private Sub alureti_Click()
'I = 0
'PASSW.Show 1
'If I = 1 Then
'    RETIRADOS.Show
'End If
End Sub

Private Sub arepogr_Click()
I = 0
PASSW.Show 1
If I = 1 Then
AREAS_GRADO.Show
End If
End Sub

Private Sub ayuedul_Click()
If Dir(App.Path & "\edulogros.hlp") = "" Then
MsgBox "NO EXISTE EL ARCHIVO DE AYUDA EDULOGROS.HLP", 48
Else
Shell "Winhlp32.exe " & App.Path & "\edulogros.hlp", vbNormalFocus
End If
End Sub

Private Sub ayulinea_Click()
If Dir("C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE") = "" Then
    MsgBox "EL NAVEGADOR WEB NO PUEDE INICIARSE, VERIFIQUE SU CONFIGURACIÓN", 48
Else
    If Dir(Ruta & "webhelp.txt") = "" Then
        MsgBox "NO EXISTE EL ARCHIVO DE LA PAGINA WEB", 64
    Else
        NAR = FreeFile
        Open Ruta & "webhelp.txt" For Input As #NAR
        Input #NAR, TTT
        Close #NAR
        Shell "C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE " & TTT, vbNormalFocus
    End If
End If
End Sub

Private Sub ayurap_Click()
HELP.Show 1
End Sub

Private Sub badispro_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    DriveBajar.Show
End If
End Sub

Private Sub bb_Click()
IMP_GRUP.Show
End Sub

Private Sub bloquedesemp_Click()
I = 0
PASSW.Show 1
If I = 1 Then
ControlPeriodos.Show 1
End If
End Sub

Private Sub cambibase_Click()
I = 0
PASSW.Show 1
If I = 1 Then
Cambiar_Base.Show 1
End If
End Sub

'Private Sub carne_Click()
'I = 0
'PASSW.Show 1
'If I = 1 Then
'CARNET.Show
'End If
'End Sub

Private Sub CarnWord_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CorrCarnet.Show 1
End If
End Sub

Private Sub cc_Click()
I = 0
PASSW.Show 1
If I = 1 Then
COPEGA.Show 1
End If
End Sub

Private Sub ccc_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CONS_NOTA.Show
End If
End Sub

Private Sub cocoalu_Click()
I = 0
PASSW.Show 1
If I = 1 Then
BASE_ALUM.Show
BASE_ALUM.Text10.SetFocus
End If
End Sub

Private Sub cocoproro_Click()
I = 0
PASSW.Show 1
If I = 1 Then
BASE_PROF.Show
BASE_PROF.Text9.SetFocus
End If
End Sub

Private Sub conalob_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CONS_NOTA.Show
End If
End Sub

Private Sub commdesempp_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    ComentariosDesemp.Show 1
End If
End Sub

Private Sub conare_Click()
CONS_MATER.Show
End Sub

Private Sub configfolio_Click()
I = 0
PASSW.Show 1
If I = 1 Then
Folio_Config.Show 1
End If
End Sub

Private Sub congru_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    ARBOL.Show
End If
End Sub

Private Sub conpro_Click()
I = 0
PASSW.Show 1
If I = 1 Then
BASE_PROF.Show
BASE_PROF.Text9.SetFocus
End If
End Sub

Private Sub cons_Click()
I = 0
PASSW.Show 1
If I = 1 Then
BASE_ALUM.Show
BASE_ALUM.Text10.SetFocus
End If
End Sub

Private Sub controlcerradas_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    Control_Planillas.Show
End If
End Sub

Private Sub Corr_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CORR_ALUM.Show 1
End If
End Sub

Private Sub corrare_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CORR_MATER.Show 1
End If
End Sub

Private Sub correpro_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CORR_PRO.Show 1
End If
End Sub

Private Sub creare_Click()
I = 0
PASSW.Show 1
If I = 1 Then
MATERIAS.Show
MATERIAS.Text1.SetFocus
End If
End Sub

Private Sub crearuser_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CrearUsers.Show 1
End If
End Sub

Private Sub cregru_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CURSOS.Show
End If
End Sub

Private Sub cvaing_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CVAAINGR.Show 1
End If
End Sub

Private Sub cveda_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CVEDAD.Show 1
End If
End Sub

Private Sub cvnoap_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CVNOMAPE.Show 1
End If
End Sub

Private Sub cvraca_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CVRANGO.Show 1
End If
End Sub

Private Sub cvrh_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CVARH.Show 1
End If
End Sub

Private Sub cvsx_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CVSEXO.Show 1
End If
End Sub

Private Sub direc_Click()
GRUPO_GRA.Show
End Sub

Private Sub directelefo_Click()
I = 0
PASSW.Show 1
If I = 1 Then
DIRECT_TEL.Show
End If
End Sub

Private Sub disdaprofe_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    DriveCopiar.Show
End If
End Sub

Private Sub eje_tem_cont_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    Ejes_Contenidos.Show
End If
End Sub

Private Sub Enlace_comp_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    Competencias.Show
End If
End Sub

Private Sub Enlace_Proyectos_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    Control_Proyectos.Show 1
End If
End Sub

Private Sub Form_Activate()
stb.Panels.Item(1).Text = "Si necesita ayuda, presione la tecla F1."
stb.Panels.Item(1).ToolTipText = stb.Panels.Item(1).Text
stb.Panels.Item(2).Text = Format(Date, "mmm/dd/yyyy")
End Sub

Private Sub Form_Deactivate()
stb.Panels.Item(1).ToolTipText = ""
End Sub

Private Sub Form_Load()
'Dim ini As inicio
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR

If Dir(Ruta & "VarEdu.edu") = "" Then
        vini.VRector = "Rector"
        vini.VDirector = "Director(a) de grupo"
        vini.VEstudiante = "Estudiante"
        vini.VGrupo = "Grupo"
        vini.VFecha = "Fecha"
        vini.VPeriodo = "Periodo"
        vini.VOp1 = ""
        vini.VOp2 = ""
        vini.VOp3 = ""
Else
    Open Ruta & "VarEdu.edu" For Input As #NAR
    Input #NAR, vini.VRector, vini.VDirector, vini.VEstudiante, vini.VGrupo, vini.VFecha, vini.VPeriodo, vini.VOp1, vini.VOp2, vini.VOp3
    Close #NAR
End If
ENTRADA.Caption = "EDULOGROS - " & ini.nombre
End Sub
Private Sub form_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
ENTRADA.PopupMenu vari
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RESP = MsgBox("DESEA SALIR DE EDULOGROS?", vbYesNo + vbQuestion + vbDefaultButton1, "SALIR")
If RESP = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub ggg_Click()
CONS_MATER.Show
End Sub

Private Sub graba_PtjLgr_Click()
If Dir(Ruta & "conf_logro.edu") <> "" Then
    Open Ruta & "conf_logro.edu" For Input As #NAR
    Input #NAR, ConfLgr
    Close #NAR
    If ConfLgr = 0 Then
        MsgBox "No se pueden grabar porcentajes de logros.  El sistema está configurado para obtener los porcentajes de forma automática.", 64, "ADVERTENCIA"
        Exit Sub
    End If
Else
    MsgBox "No se pueden grabar porcentajes de logros.  El sistema está configurado para obtener los porcentajes de forma automática.", 64, "ADVERTENCIA"
    Exit Sub
End If
I = 0
PASSW.Show 1
If I = 1 Then
Porcentaje_Logros.Show
End If
End Sub

Private Sub grabadesemp_Click()
I = 0
PASSW.Show 1
If I = 1 Then
GRABA_DESEMP.Show
End If
End Sub

Private Sub grabogru_Click()
I = 0
PASSW.Show 1
If I = 1 Then
GRABAR_OBSER.Show
End If
End Sub

Private Sub grabdesemp2_Click()
I = 0
PASSW.Show 1
If I = 1 Then
GRABA_DESEMP.Show
End If
End Sub

Private Sub gralu_Click()
GRADOS.Show
End Sub

Private Sub grumodi_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    CONS_GRUP.Show
End If
End Sub

Private Sub histo_Click()
I = 0
PASSW.Show 1
If I = 1 Then
HISTORIAL.Show 1
End If
End Sub

Private Sub imalob_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CONS_NOTA.Show
End If
End Sub

Private Sub impr_Click()
IMP_GRUP.Show
End Sub

Private Sub infoadi_Click()
I = 0
PASSW.Show 1
If I = 1 Then
info_adicional.Show 1
End If
End Sub

Private Sub infoareas_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    NOTA_GENRL.Show
End If
End Sub

Private Sub infoareas2_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    REPORTE_PORCENT.Show
End If
End Sub

Private Sub infofinal_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    'Dim leye As leyendis
    If Dir(Ruta & "leyenda.edu") <> "" Then
        NAR = FreeFile
        Open Ruta & "leyenda.edu" For Input As #NAR
        Input #NAR, leye.ly1, leye.ly2, leye.ly3, leye.ly4, leye.ly5, leye.ly6, leye.ly7, leye.ly8
        Close #NAR
        LEYENDA.Text1.Text = RTrim(leye.ly1)
        LEYENDA.Text2.Text = RTrim(leye.ly2)
        LEYENDA.Text3.Text = RTrim(leye.ly3)
        LEYENDA.Text4.Text = RTrim(leye.ly4)
        LEYENDA.Text5.Text = RTrim(leye.ly5)
        LEYENDA.Text6.Text = RTrim(leye.ly6)
        LEYENDA.Text7.Text = RTrim(leye.ly7)
        LEYENDA.Text8.Text = RTrim(leye.ly8)
    Else
        LEYENDA.Text1.Text = ""
        LEYENDA.Text2.Text = ""
        LEYENDA.Text3.Text = ""
        LEYENDA.Text4.Text = ""
        LEYENDA.Text5.Text = ""
        LEYENDA.Text6.Text = ""
        LEYENDA.Text7.Text = ""
        LEYENDA.Text8.Text = ""
    End If
    LEYENDA.Show 1
End If
End Sub

Private Sub infrfin_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    INFREFINAL.Show
End If
End Sub

Private Sub inian_Click()
I = 0
PASSW.Show 1
If I = 1 Then
  RESP = MsgBox("1. DESEA CREAR EL HISTORIAL? (SI YA LO CREO DE CLICK EN NO)", vbYesNo + vbQuestion + vbDefaultButton1, "ADVERTENCIA")
  If RESP = vbYes Then
     HISTORIAL.Show 1
     Else
     INI_ANO.Show 1
  End If
End If
End Sub

Private Sub inici_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    DAT_INI.Show 1
End If
End Sub

Private Sub izconsugrup_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    ARBOL.Show
End If
End Sub

Private Sub logporperiod_Click()
PEND_GENRL.Show
End Sub

Private Sub logporperiod2_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    Reporte_Logros.Show
End If
End Sub

Private Sub menparapro_Click()
I = 0
PASSW.Show 1
If I = 1 Then
PARAPROMO.Show 1
End If
End Sub

Private Sub mit_impr_Click()
I = 0
PASSW.Show 1
If I = 1 Then
cons_nota2.Show
End If
End Sub

Private Sub mit_logreap_Click()
PEND_GENRL2.Show
End Sub

Private Sub obse_Click()
I = 0
PASSW.Show 1
If I = 1 Then
COPEGA.Show 1
End If
End Sub

Private Sub Orargru_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    Ord_argru.Show 1
End If
End Sub

Private Sub paradesemp_Click()
I = 0
PASSW.Show 1
If I = 1 Then
CONFIG_DESEMP.Show 1
End If
End Sub

Private Sub pensi_Click()
I = 0
PASSW.Show 1
If I = 1 Then
PENSIONES.Show 1
End If
End Sub

Private Sub PRORETI_Click()
'I = 0
'PASSW.Show 1
'If I = 1 Then
'    RETI_PRO.Show
'    RETI_PRO.Combo1.SetFocus
'End If
End Sub

Private Sub planea_sema_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    planeacion_semanal.Show
End If
End Sub

Private Sub porlogros44_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    MsgBox "Tenga en cuenta que el cambio de configuración de porcentajes de logros afecta la forma como se muestran los resultados en el sistema y la forma como ingresan las notas los profesores", 48
    Conf_Logros.Show 1
End If
End Sub

Private Sub refinal_Click()
I = 0
PASSW.Show 1
If I = 1 Then
RESUFINA.Show
End If
End Sub

Private Sub regdataprof_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    If Dir(Ruta & "infnota.edu") = "" Then
        MsgBox "No hay información disponible", 64, "Verificar"
        Exit Sub
    End If
    Screen.MousePointer = 11
    VERI_PROSIS.MATI28.ColWidth(0) = 4000
    VERI_PROSIS.MATI28.ColWidth(1) = 1000
    VERI_PROSIS.MATI28.ColWidth(2) = 1000
    VERI_PROSIS.MATI28.ColWidth(3) = 1200
    VERI_PROSIS.MATI28.Row = 0
    VERI_PROSIS.MATI28.Col = 0
    VERI_PROSIS.MATI28.CellForeColor = RGB(255, 255, 255)
    VERI_PROSIS.MATI28.CellBackColor = RGB(0, 0, 150)
    VERI_PROSIS.MATI28.Text = "NOMBRE DEL PROFESOR"
    VERI_PROSIS.MATI28.Col = 1
    VERI_PROSIS.MATI28.CellForeColor = RGB(255, 255, 255)
    VERI_PROSIS.MATI28.CellBackColor = RGB(0, 0, 150)
    VERI_PROSIS.MATI28.Text = "PERIODO"
    VERI_PROSIS.MATI28.Col = 2
    VERI_PROSIS.MATI28.CellForeColor = RGB(255, 255, 255)
    VERI_PROSIS.MATI28.CellBackColor = RGB(0, 0, 150)
    VERI_PROSIS.MATI28.Text = "FECHA"
    VERI_PROSIS.MATI28.Col = 3
    VERI_PROSIS.MATI28.CellForeColor = RGB(255, 255, 255)
    VERI_PROSIS.MATI28.CellBackColor = RGB(0, 0, 150)
    VERI_PROSIS.MATI28.Text = "HORA"
    J = 1
    NAR = FreeFile
    Open Ruta & "infnota.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, ifnt.numprofe, ifnt.periodo, ifnt.fecha, ifnt.hora
        NAR = FreeFile
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, ifnt.numprofe, profe
        Close #NAR
        VERI_PROSIS.MATI28.Rows = J + 1
        VERI_PROSIS.MATI28.TextMatrix(J, 0) = RTrim(profe.nombres) & " " & RTrim(profe.apellidos) & "(" & ifnt.numprofe & ")"
        VERI_PROSIS.MATI28.TextMatrix(J, 1) = ifnt.periodo
        VERI_PROSIS.MATI28.TextMatrix(J, 2) = ifnt.fecha
        VERI_PROSIS.MATI28.TextMatrix(J, 3) = ifnt.hora
        J = J + 1
        NAR = NAR - 1
    Wend
    Close #NAR
    Screen.MousePointer = 0
    VERI_PROSIS.Show 1
End If

End Sub

Private Sub reprecu_Click()
'I = 0
'PASSW.Show 1
'If I = 1 Then
'IMP_RECU.Show 1
'End If
End Sub

Private Sub retalma_Click()
I = 0
PASSW.Show 1
If I = 1 Then
RETI_ALUM.Show
End If
End Sub

Private Sub retipro_Click()
I = 0
PASSW.Show 1
If I = 1 Then
RETIS.Show
End If
End Sub

Private Sub salis_Click()
End
End Sub

Private Sub singrupen_Click()
PENGRU.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "datalu"
I = 0
PASSW.Show 1
If I = 1 Then
BASE_ALUM.Show
BASE_ALUM.Text10.SetFocus
End If
Case "datapro"
I = 0
PASSW.Show 1
If I = 1 Then
BASE_PROF.Show
BASE_PROF.Text9.SetFocus
End If
Case "matric"
I = 0
PASSW.Show 1
If I = 1 Then
MATRICULA.Show 1
End If
Case "alumexi"
GRADOS.Show
Case "directele"
I = 0
PASSW.Show 1
If I = 1 Then
DIRECT_TEL.Show
End If
Case "congru"
I = 0
PASSW.Show 1
If I = 1 Then
ARBOL.Show
End If
Case "impgru"
IMP_GRUP.Show
Case "contotal"
TOTALES.Show
Case "conare"
CONS_MATER.Show
Case "adiobs"
I = 0
PASSW.Show 1
If I = 1 Then
COPEGA.Show 1
End If
Case "aregra"
I = 0
PASSW.Show 1
If I = 1 Then
AREAS_GRADO.Show
End If
Case "agreobs"
I = 0
PASSW.Show 1
If I = 1 Then
GRABAR_OBSER.Show
End If
Case "adiobal"
I = 0
PASSW.Show 1
If I = 1 Then
'GRABAR_OBSER.Show
GRABA_DESEMP.Show
End If
Case "conobal"
I = 0
PASSW.Show 1
If I = 1 Then
CONS_NOTA.Show
End If
Case "impobal"
I = 0
PASSW.Show 1
If I = 1 Then
CrearUsers.Show 1
End If
Case "esta"
PEND_GENRL.Show
Case "disdapro"
I = 0
PASSW.Show 1
If I = 1 Then
DriveCopiar.Show
End If
End Select
End Sub

Private Sub tota_Click()
TOTALES.Show
End Sub

Private Sub txt_encab2_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    conf_encabeza2.Show 1
End If
End Sub
