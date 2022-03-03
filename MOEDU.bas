Attribute VB_Name = "MOEDU"
Option Explicit
Public I As Integer, r As Integer, h As Integer, HO As Integer, w As Integer, ww As Integer, JO As String, AI As String, CER As String, k As Integer, VERI As Integer, AT As String, CONT As Integer, TT As Integer, VALI As Boolean, VALI2 As Boolean, VALI4 As Boolean, leo As Integer, curcar As Integer, LEO2 As Integer, plo As Integer, rr As Integer, falta As Integer, cort As String, CERD As Integer, CH As Integer, CM As Integer, SP As Integer, PEGG As String, VVAA As Boolean
Public J As Integer, z As Integer, Y As Integer, s As Integer, t As Integer, VERIFI As Boolean, VERIFI2 As Boolean, PAG As Integer, dire As Integer, dire2 As Integer, GH As Integer, J1 As String, J2 As String, J4 As String, J5 As String, PARCHI As String, YUS As String, ar As String, gua As Integer, YO As Integer, VALI44 As Boolean, OPP As Integer, CLIS As Integer, TN As Integer, TTT As String, rt As Integer, trt As Currency, spa As Currency, SINCON As Boolean
Public que As Integer, maa As Integer, cli As Integer, BOL As Boolean, NA As Integer, PERI As String, VALI80 As Boolean, VALI180 As Boolean, VALI380 As Boolean, cona As Integer, cona2 As Integer, CLO As Integer, CHA As Integer, CROA As Integer, Ver_Ini As Integer, ABC As Boolean, zo As Integer, zi As Integer, ki As Integer, clat As Integer, sir As Integer, rei As Integer, SIRO As Integer, cruz As Integer, RESC As String, we As Integer, PENTI As Boolean, tpm As Currency, tgp As Currency, L As Integer, IMPOK As Boolean
Public pio As Integer, p As Integer, q As Integer, ret As Integer, SUR As Integer, VV As Integer, ape As String, nom As String, FERT As Integer, RE22 As String, REE22 As String, ja As Integer, lw As Integer, ser As String, jho As String, zu As Integer, term As Integer, NJOR As String, NGRA As String, OB As String, INDI As String, año As String, cop As String, clacerr As Boolean, grune As Boolean, d As Integer, nu As Integer, nf As Integer, EXISALU As Boolean
Public RECO As Boolean, CED As String, QQ As Integer, CUCU As Integer, DF As Integer, seri As String, lwe As Integer, fl As String, cru As Integer, zz As Integer, uu As Integer, rus As Integer, fis As Integer, CLA As Integer, malo As Integer, CONTAREA As Integer, noyu As Integer, JOJI As String, SAPO As String, SAPO2 As String, SAPO3 As String, SAPO4 As String, CX As Single, CY As Single, NAR As Integer, J3 As Integer, MS1 As String, CP As Integer, AdiCampo As newcampo
Public alumno As maestroalum, profe As maestropro, icur As inforcur, ifnt As infornoti, mate As infomater, alugru As grupoalu, aluper As pertgrup, hisalugru As hisgrupoalu, argra As areagr, logru As logris, retiros As retiro, notas As notis, hisnotas As hisnotis, proti As pro_reti, ini As inicio, comdpe As comdesemp, vini As VarInicio, contra As clave, CLAV As CLAVEPRO, leye As leyendis, leyfin As leyenfin, hisleyfin As hisleyenfin, obsfin As String * 750, pens(1 To 12) As Currency, infsub As inforsub
Public YF As Single, YI As Single, YC As Single, VR As String, CoB As Byte, NObs As Boolean, OkArea As Boolean, newmatri As infomatri, Ruta As String, RutaDir As String, RutaCSV As String, detalle As info_detalle
Public confdesemp As ini_desemp, notas_desemp As porcentaje_desemp, SWobserv As Boolean, Cont_Lgr As Integer, proflogs As bitacora, ConfTexto As String, porcent_manual As porcentaje_manual
Public proyectosaula As proyectos, semanal_planeacion As planeacion_semanal, semanal_ejetematico As eje_tematico_semanal, semanal_contenidos As contenidos_semanal, semanal_competencias As competencias_semanal
Public ValiModifica As Boolean, Mod_Vr As String, NewPyArea As Boolean


Type maestroalum
n_matricula As Integer
n_carnet As String * 8
nombres As String * 30
apellidos As String * 30
documento As String * 15
f_nacimiento As String * 10
rh As String * 4
sexo As String * 1
padre As String * 50
tel_pa As String * 20
madre As String * 50
tel_ma As String * 20
acudiente As String * 50
tel_acu As String * 20
direccion As String * 60
jornada As String * 10
año_ingre As String * 4
grado As String * 15
End Type

Type maestropro
nombres As String * 30
apellidos As String * 30
documento As String * 15
fech_nacim As String * 10
rh As String * 4
direccion As String * 60
Telefono As String * 20
año_ingre As String * 4
especiali As String * 40
escalafon As String * 2
End Type

Type inforcur
nom As String
jornada As String
grado As String
director As Integer
End Type

Type infornoti
numprofe As Integer
periodo As String
fecha As Date
hora As String
End Type

Type inforsub
subsistema As Integer
bajasub As String
actualsub As String
End Type

Type infomater
nom As String * 50
num As Integer
End Type

Type grupoalu
num_carnet As String * 5
End Type

Type pertgrup
grupo As String * 20
End Type

Type newcampo
salud As String * 40
Tel_casa As String * 20
email As String * 50
otras As String * 100
End Type

Type infomatri
nombre(1 To 14) As String * 30
grado(1 To 14) As String * 15
año(1 To 14) As String * 10
ciudad(1 To 14) As String * 10
End Type

Type hisgrupoalu
nombres As String * 30
apellidos As String * 30
End Type

Type areagr
grado As String * 15
num_area As Integer
ih As Integer
num_pro As Integer
nom_grup As String * 20
End Type

Type logris
indicador As String * 5
observ As String * 800
End Type

Type retiro
nombres As String * 30
apellidos As String * 30
direccion As String * 60
Telefono As String * 20
jornada As String * 10
año_ingreso As String * 4
año_retiro As String * 4
grado As String * 15
End Type

Type notis
num_carnet As String * 5
FA As Integer
area(1 To 10) As Integer
End Type

Type hisnotis
nombres As String * 30
apellidos As String * 30
JV As String * 2
FA As Integer
nota(1 To 10) As Integer
End Type

Type pro_reti
nombres As String * 30
apellidos As String * 30
documento As String * 15
rh As String * 4
direccion As String * 60
Telefono As String * 20
año_ingre As String * 4
año_retir As String * 4
especiali As String * 40
escalafon As String * 2
End Type

Type inicio
ciudad As String    'SE CAMBIO PARA MOSTRAR LA RESOLUCION
nombre As String
modalidad As String 'SE CAMBIO PARA MOSTRAR INFORMACION OPCIONAL EN EL ENCABEZADO DEL REPORTE
Telefono As String  'SE CAMBIO PARA MOSTRAR EL AÑO ACADEMICO
Rector As String
secretario As String
End Type

Type comdesemp
bajo As String
basico As String
alto As String
superior As String
End Type

Type VarInicio
VRector As String
VDirector As String
VEstudiante As String
VGrupo As String
VFecha As String
VPeriodo As String
'Variable opcional 1
VOp1 As String
'Variable opcional 2
VOp2 As String
'Variable opcional 3
VOp3 As String
End Type

Type clave
nombre As String * 150
PASSW As String * 150
End Type

Type CLAVEPRO
NUMERO As Integer
PASSWW As String * 15
End Type

Type leyendis
ly1 As String
ly2 As String
ly3 As String
ly4 As String
ly5 As String
ly6 As String
ly7 As String
ly8 As String
End Type

Type leyenfin
num_carnet As String * 5
fnob(1 To 5) As Integer
End Type

Type hisleyenfin
nombres As String * 30
apellidos As String * 30
fnob(1 To 5) As Integer
End Type

Type info_detalle
info As String * 2000
End Type

Type bitacora
numprofe As Integer
fecha As Date
hora As String
End Type

Type ini_desemp
grado As String * 15
desemp(1 To 4) As String * 5
recupera(1 To 4) As String * 5
rango(1 To 3) As Byte
End Type

Type porcentaje_desemp
num_carnet As String * 5
porcentaje(1 To 10) As Byte
recuperado(1 To 10) As Boolean
logro(1 To 10) As Byte
End Type

'PORCENTAJE MANUAL DE LOGROS
Type porcentaje_manual
porcent_logro As Byte
End Type

'**********************************
'ESTRUCTURA DE DATOS -PLANEACIONES-
'**********************************

' PLANEACION SEMANAL

Type planeacion_semanal
fecha As String * 20
eje As String * 100
contenidos As String * 100
competencia As String * 100
logros As String * 100
End Type

Type eje_tematico_semanal
'num_eje As Integer
txt_eje As String * 200
End Type

Type contenidos_semanal
'num_cont As Integer
txt_cont As String * 200
num_eje As Integer
End Type

Type competencias_semanal
'num_comp As Integer
cod_comp As String * 10
txt_comp As String * 700
num_logro As String * 50
End Type

'Type logro_competencia_semanal
'num_comp As Integer
'num_logro As Integer
'End Type

' PROYECTOS

Type proyectos
nombre As String * 500
responsables As String * 1000
poblacion As String * 500
objetivos As String * 2000
Competencias As String * 5000
metas As String * 2000
ejes_tematicos As String * 2000
metodologia As String * 2000
cronograma As String * 2000
recursos As String * 2000
evaluacion As String * 2000
observaciones As String * 2000
End Type
