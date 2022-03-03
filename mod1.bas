Attribute VB_Name = "MOEDU"
Option Explicit
Public i As Integer, r As Integer, h As Integer, HO As Integer, w As Integer, ww As Integer, JO As String, AI As String, CER As String, k As Integer, VERI As Integer, AT As String, co As Integer, CONT As Integer, TT As Integer, VALI1 As Boolean, VALI2 As Boolean, VALI3 As Boolean, VALI4 As Boolean, KK As Integer, cc As Integer, FF As Integer, leo As Integer, plo As Integer, rr As Integer, falta As Integer, TER As Integer, TER2 As Integer, re5 As Integer, crelogro2 As String, crelogro As String, cort As String, flar As Integer, CERD As Integer, pty As Integer, CH As Integer, CM As Integer, SP As Integer
Public J As Integer, BORR As Integer, z As Integer, y As Integer, s As Integer, t As Integer, ape1 As String, ape2 As String, nom1 As String, nom2 As String, fec1 As String, fec2 As String, VERIFI As Boolean, VERIFI2 As Boolean, PAG As Integer, dire As Integer, GH As Integer, J1 As String, J2 As String, JJ As Integer, PARCHI As String, T5 As Integer, LL As Integer, YUS As String, ar As String, gua As Integer, YO As Integer, VALI44 As Boolean, NIM As Integer, OPP As Integer, CLIS As Integer, TN As Integer, YUR As Integer, lio As Integer, PIG As Integer, TTT As String, rt As Integer, trt As Long, spa As Long
Public que As Integer, maa As Integer, cli As Integer, BOL As Boolean, NP As Integer, NA As Integer, PERI As String, VALI80 As Boolean, cona As Integer, cona2 As Integer, VVAA As Boolean, AABB As Boolean, CLO As Integer, CHA As Integer, CROA As Integer, ABC As Boolean, zo As Integer, zi As Integer, rett As Boolean, ki As Integer, clat As Integer, sir As Integer, rei As Integer, SIRO As Integer, cruz As Integer, RESC As String, we As Integer, ddi As Integer, VACA As Integer, PENTI As Boolean, COOT As Integer, HUT As Integer, HIR As Integer, RES As Integer, tpm As Long, tgp As Long, L As Integer, AÑO As String
Public pio As Integer, p As Integer, q As Integer, ret As Integer, SUR As Integer, VV As Integer, ert As Integer, matri As Integer, ape As String, nom As String, FERT As Integer, FERT2 As Integer, RE11 As String, RE22 As String, REE22 As String, ABRIR As String, NAB As Integer, CLAS As Integer, QW As Integer, MOU As Integer, HOY As Integer, clare As Integer, HOY2 As Integer, jur As Integer, kur As Integer, ja As Integer, lw As Integer, lw2 As Integer, ser As String, ser2 As String, YY As Integer, tito As Integer, jho As String, zu As Integer, term As Integer, ris As Integer, NJOR As String, NGRA As String, copiarch As String
Public RECO As Boolean, CED As String, QQ As Integer, CUCU As Integer, DF As Integer, re32 As String, seri As String, noar As String, lwe As Integer, fl As Integer, cru As Integer, tensi As Integer, pp As Integer, zz As Integer, uu As Integer, xo As Integer, yi As Integer, rus As Integer, fis As Integer, CLAV As Integer, malo As Integer, giti As Integer, gli As Integer, CONTAREA As Integer, jis As Integer, noyu As Integer, ver As Integer, JOJI As String, SAPO As String, SAPO2 As String, SAPO3 As String, TOS As Integer, CX As Single, CY As Single, NAR As Integer, DISP As Integer, J3 As Integer, MS1 As String

Type maestroalum
n_matricula As Integer
n_carnet As String * 8
nombres As String * 20
apellidos As String * 20
documento As String * 12
f_nacimiento As String * 10
rh As String * 4
sexo As String * 1
padre As String * 30
tel_pa As String * 12
madre As String * 30
tel_ma As String * 12
acudiente As String * 30
tel_acu As String * 12
direccion As String * 40
jornada As String * 6
año_ingre As String * 4
grado As String * 10
End Type

Type maestropro
nombres As String * 20
apellidos As String * 20
documento As String * 10
fech_nacim As String * 10
rh As String * 4
direccion As String * 40
telefono As String * 12
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
fecha As Variant
End Type

Type infomater
nom As String * 23
num As Integer
End Type

Type grupoalu
num_carnet As String * 5
End Type

Type hisgrupoalu
nombres As String * 20
apellidos As String * 20
End Type

Type areagr
grado As String * 10
num_area As Integer
ih As Integer
num_pro As Integer
nom_grup As String * 13
End Type

Type logris
indicador As String * 1
observ As String * 150
End Type

Type retiro
nombres As String * 20
apellidos As String * 20
direccion As String * 40
telefono As String * 12
jornada As String * 6
año_ingreso As String * 4
año_retiro As String * 4
grado As String * 10
End Type

Type notis
num_carnet As String * 5
JV As String * 2
FA As Integer
area(1 To 10) As Integer
End Type

Type hisnotis
nombres As String * 20
apellidos As String * 20
JV As String * 2
FA As Integer
nota(1 To 10) As Integer
End Type

Type pro_reti
nombres As String * 20
apellidos As String * 20
documento As String * 10
rh As String * 4
direccion As String * 40
telefono As String * 12
año_ingre As String * 4
año_retir As String * 4
especiali As String * 40
escalafon As String * 2
End Type

Type inicio
ciudad As String
nombre As String
modalidad As String
telefono As String
End Type

Type clave
nombre As String * 15
PASSW As String * 15
End Type

Type CLAVEPRO
NUMERO As Integer
PASSWW As String * 15
End Type

Type pension
pe(1 To 12) As Long
grado As String * 10
jornada As String * 6
End Type
