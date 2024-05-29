Attribute VB_Name = "Samag"
Dim rgGuias As Range, rgToolsSamag As Range
Dim opIndex As Integer, lbl As Integer, guia As Integer

Sub SamagInitialize()
  Set rgGuias = [nmGuias]
  rgGuias(2, 1) = 11
  rgGuias(2, 2) = 12
  rgGuias(2, 3) = 13
  rgGuias(2, 4) = 14
End Sub

Sub SamagMain()
'
Dim avOpName As Variant, avOpShortName As Variant

If fnValidaDados(2) = False Then Exit Sub

Set rgGuias = [nmGuias]
'Exit se células vazias na tabela de guias


Set rgToolsSamag = [toolsSamag]

avOpName = Array("Desbaste", "Pre", "Fundo", "Escarear", "Acaba") 'para nome de main

For opIndex = 1 To 5
  'nomeFicheiro = "Guia_" & opIndex & avOpName(opIndex - 1) & ".h"
  nomeFicheiro = "Guia_" & opIndex
  f = Sheets("cfg").[A1] & "\" & nomeFicheiro & ".h"
  Open f For Output As #1
  
  '-------- Header ------------------
  Print #1, "BEGIN PGM " & nomeFicheiro & " MM"
  
  pgmHeader = ";Guias Samag - " & avOpName(opIndex - 1)
  'Debug.Print fnSeparador(pgmHeader)
  Print #1, pgmHeader
  'Debug.Print fnSeparador(pgmHeader)
   
  '-------- Body ------------------
  If opIndex <> 5 Then
    ExceptoAcaba 'todas as operações excepto "Acaba"
    Print #1, "LBL " & Chr(34) & "Fim" & Chr(34)
  Else
    Acaba
  End If

  '-------- Footer ------------------
  Print #1, "END PGM " & nomeFicheiro & " MM"
  Close #1
Next
  MsgBox "ok"
End Sub

Sub ExceptoAcaba()

'tool
sToolCall = "TOOL CALL " & rgToolsSamag(1, opIndex) & " Z" & _
  " S" & rgToolsSamag(2, opIndex)
  
Print #1, sToolCall

'-------- Inputs ------------------
'set opIndex (definido aqui apenas para opIndex 1,3 porque partilham sub
If opIndex = 1 Or opIndex = 3 Then
  Print #1, "Q20=" & opIndex & " ;op index (NAO ALTERAR)"
End If

Print #1, "Q21=50 ;dZ seguranca" 'dZsecurity

'dyDesbastado
If opIndex = 1 Then
  Print #1, "Q22=60 ;dY desbastado" 'dyDesbastado
ElseIf opIndex = 3 Then
  Print #1, "Q22=0 ;dY desbastado" 'dyDesbastado
End If

'pré: corte num ou dois sentidos (Q29)
If opIndex = 2 Then
  Print #1, "Q29=1 ;Corte nos dois sentidos"
End If

'loop para CALL LBL
For guia = 1 To 4
  If rgGuias(2, guia) <> "" Then
    Print #1, "CALL LBL " & rgGuias(2, guia)
  End If
Next
Print #1, "FN 9: IF +1 EQU +1 GOTO LBL " & Chr(34) & "Fim" & Chr(34)

'-------------------------- inputs para cada guia ----------------------------
'nome da sub (chamada)
avSubName = Array("1", "2", "1", "4", "5") 'porque opIndex 1 e 3 partilham sub
subName = "GuiaSub" & avSubName(opIndex - 1) & ".h"
zFace = 0
For guia = 1 To 4
  If rgGuias(2, guia) <> "" Then 'guia TRUE
    Print #1, ";----"
    Print #1, "LBL " & rgGuias(2, guia) & " ; B" & rgGuias(1, guia)
    
    Print #1, "Q0=" & rgGuias(2, guia) & " ;origem"
    Print #1, "Q1=" & rgGuias(3, guia) & " ;dimX"
    Print #1, "Q2=" & rgGuias(4, guia) & " ;dimY"
    'dimZ: existe em todos excepto escarear
    If opIndex <> 4 Then Print #1, "Q3=" & rgGuias(5, guia) & " ;dimZ"
    'zFace: existe em todos
    Print #1, "Q4=" & zFace & " ;zFace"

    'zInicial: só escreve se opIndex=1,2,3
    If opIndex < 4 Then
      Select Case opIndex
        Case 1, 2
          zInicial = zFace 'default=zFace
        Case 3
          'default=zFace-dimZ+stock do desbaste
            'zInicial = zFace - rgGuias(5, guia) + rgTecno(8, 1)
          'passámos a fazer apenas a passagem a zFinal
          zInicial = zFace - rgGuias(5, guia)
      End Select
      
      Print #1, "Q5=" & zInicial; " ;zInicial"
    End If
    
    'varar
    Print #1, "Q6=" & rgGuias(6, guia) & " ;varar 0=nao"
    'chamada da sub, excepto para acaba
    Print #1, "CALL PGM TNC:\tcm\guia\" & subName
    
    Print #1, "LBL 0"
  End If
Next
'
End Sub

Sub Acaba()
'
Print #1, "Q0 = 11 ;Id guia"
Print #1, ";settings acaba"
Print #1, "Q21=50 ;dZ seguranca" 'dZsecurity
Print #1, "Q23=0.3 ;stock inicial"
Print #1, "Q24=1 ;repete ultima passagem (1=Sim 0=Nao)"
Print #1, "Q26 = 0 ;Descentramento X"
Print #1, ";settings recentra"
Print #1, "Q81 = 0 ;Set Z"
Print #1, "Q82 = 1 ;Posiciona B"
Print #1, "Q83 = 20 ;dZ seguranca"
Print #1, "Q84 = -5 ;Z medir X"
Print #1, "Q85 = -13 ;Y para centrar X"
Print #1, ";---------------------"
Print #1, "CALL LBL Q0"
'recentra
'Print #1, "CALL LBL " & Chr(34) & "Recentra" & Chr(34)
Print #1, "CALL PGM TNC:\tcm\guia\recentra.h"
'tool
sToolCall = "TOOL CALL " & rgToolsSamag(1, opIndex) & " Z" & _
  " S" & rgToolsSamag(2, opIndex)

Print #1, sToolCall
Print #1, "CALL PGM TNC:\tcm\guia\GuiaSub5.h"
'mede
Print #1, "/CALL PGM TNC:\tcm\guia\mede.h"
Print #1, "M30"

'-------------------------- inputs para cada guia ----------------------------
'nome da sub (chamada)
avSubName = Array("1", "2", "1", "4", "5") 'porque opIndex 1 e 3 partilham sub
subName = "GuiaSub" & avSubName(opIndex - 1) & ".h"
zFace = 0
For guia = 1 To 4
  If rgGuias(2, guia) <> "" Then 'guia TRUE
    Print #1, ";----"
    Print #1, "LBL " & rgGuias(2, guia) & " ; B" & rgGuias(1, guia)
        
    Print #1, "Q1=" & rgGuias(3, guia) & " ;dimX"
    Print #1, "Q2=" & rgGuias(4, guia) & " ;dimY"
    'dimZ: existe em todos excepto escarear
    Print #1, "Q3=" & rgGuias(5, guia) & " ;dimZ"
    'zFace: existe em todos
    Print #1, "Q4=" & zFace & " ;zFace"
    'varar
    Print #1, "Q6=" & rgGuias(6, guia) & " ;varar 0=nao"
    
'--- esta variável passou para o grupo de "settings recentra" (Q85)
    'Y para medir X, no ciclo para recentrar
    'Print #1, "Q7=-13 ;Y para medir X"
    
    Print #1, "LBL 0"
  End If
Next
'
End Sub

Function fnSeparador(s1)
  'um pequeno preciosismo
  For i = 1 To Len(s1) - 1
    s2 = s2 & "-"
  Next
  
  fnSeparador = ";" & s2
End Function

Sub BttPresetSamag()
  If fnValidaDados(2) = False Then Exit Sub
  
  GetPresetValues
  PresetOrigens
End Sub

Sub GetPresetValues()
'
' Esta sub pode ser integrada na 'PresetOrigens'

Set rgPreset = [nmPreset]

f = "D:\pt\out\PRESET.PR"
Open f For Input As #1
For i = 1 To 4
  Line Input #1, lin
Next
Close #1

'X
str1 = Mid(lin, 38, 12)
xPr = CDbl(Trim(str1))
rgPreset(1) = CDbl(Trim(str1))
'Y
str1 = Mid(lin, 50, 12)
yPr = CDbl(Trim(str1))
rgPreset(2) = CDbl(Trim(str1))
'Z
str1 = Mid(lin, 62, 12)
zPr = CDbl(Trim(str1))
rgPreset(3) = CDbl(Trim(str1))
'B
str1 = Mid(lin, 86, 12)
bPr = CDbl(Trim(str1))
rgPreset(4) = CDbl(Trim(str1))

MsgBox "Ok, GetPresetValues"
'
End Sub


Sub PresetOrigens()
'
Set rgGuias = [nmGuias]
Set rgPreset = [nmPreset]

'valores da origem de referência (dft O1), copiados da tabela de PRESET
'X e Z no centro, Y na base

'x1 = -0.2043
'y1 = 1016.8134
'z1 = 25.015
'b1 = 270.0033

'lê valores da folha, importados por 'GetPresetValues'
x1 = rgPreset(1)
y1 = rgPreset(2)
z1 = rgPreset(3)
b1 = rgPreset(4)

dimX = [B2]
dimZ = [B3]

f = Sheets("cfg").[A1] & "\PresetOri.h"
Open f For Output As #1

Print #1, "BEGIN PGM PresetOri MM "

'ciclo para todas as origens
For guia = 1 To 4
  If rgGuias(2, guia) <> "" Then
    ori = rgGuias(2, guia)
    'cálculo de X e Z
    rot = rgGuias(1, guia)
    beta = Application.WorksheetFunction.Radians(rot)
    
    'rotação do centro da peça (x2,z2)
    x2 = x1 * Cos(beta) - z1 * Sin(beta)
    z2 = z1 * Cos(beta) + x1 * Sin(beta)
    
    'translação para a face
    dxCentro = rgGuias(7, guia) 'desvio (ex: guia deslocada)
    'deslocamento Z (metade de dimX ou dimZ)
    Select Case rot
      Case 0, 180
        dzCentro = dimZ / 2 'dimZ/2
      Case 90, 270, -90
        dzCentro = dimX / 2 'dimX/2
    End Select

    xOri = x2 + -(dxCentro) 'escrevi desta forma para reforçar a TROCA DO SINAL
    zOri = z2 - dzCentro
    bOri = b1 - rot
    If bOri < 0 Then bOri = bOri + 360
    yOri = y1 - rgGuias(8, guia)
    
    'Debug.Print rot, x2, z2
    'Debug.Print rot, xOri, yOri, zOri, bOri
    Print #1, "FN 17: SYSWRITE ID 503 NR" & ori & " IDX1 = " & Round(xOri, 4)
    Print #1, "FN 17: SYSWRITE ID 503 NR" & ori & " IDX2 = " & Round(yOri, 4)
    Print #1, "FN 17: SYSWRITE ID 503 NR" & ori & " IDX3 = " & Round(zOri, 4)
    Print #1, "FN 17: SYSWRITE ID 503 NR" & ori & " IDX5 = " & Round(bOri, 4)
  End If
Next

Print #1, "END PGM PresetOri MM "
Close #1
MsgBox "Ok, PresetOri"
'
End Sub

Sub SamagCentraGuias()
'
If fnValidaDados(2) = False Then Exit Sub

Set rgGuias = [nmGuias]

f = Sheets("cfg").[A1] & "\CentraGuias.h"
Open f For Output As #1

Print #1, "BEGIN PGM CentraGuias MM "
Print #1, "Q3 = - 13 ;Y para medir X"
Print #1, "Q4 = - 5 ;Z para medir X"

For guia = 1 To 4
  If rgGuias(2, guia) <> "" Then
    Print #1, ";"
    Print #1, "QR0=" & rgGuias(2, guia) & " ;origem"
    Print #1, "Q1=" & rgGuias(3, guia) & " ;dimX"
    Print #1, "Q2=-" & rgGuias(5, guia) & " ;dimZ"
    Print #1, "CALL LBL 10"
  End If
Next

Print #1, "M30"
Print #1, ";"
Print #1, "LBL 10"
Print #1, "CALL PGM TNC:\tcm\OrigemB0.H"
Print #1, "TCH PROBE 408 PTO.REF.CENTRO RAN. ~"
Print #1, "    Q321=+0    ;CENTRO DO 1. EIXO ~"
Print #1, "    Q322=+Q3   ;CENTRO DO 2. EIXO ~"
Print #1, "    Q311=+Q1   ;LARGURA RANHURA ~"
Print #1, "    Q272=+1    ;EIXO DE MEDICAO ~"
Print #1, "    Q261=+Q4   ;ALTURA MEDIDA ~"
Print #1, "    Q320=+5    ;DISTANCIA SEGURANCA ~"
Print #1, "    Q260=+20   ;ALTURA DE SEGURANCA ~"
Print #1, "    Q301=+0    ;IR ALTURA SEGURANCA ~"
Print #1, "    Q305=+QR0 ;NUMERO NA TABELA ~"
Print #1, "    Q405=+0    ;PONTO DE REFERENCIA ~"
Print #1, "    Q303=+1    ;TRANSM. VALOR MED. ~"
Print #1, "    Q381=+1    ;APALPAR NO EIXO TS ~"
Print #1, "    Q382=+0    ;1. COORD. EIXO TS ~"
Print #1, "    Q383=+Q3   ;2. COORD. EIXO TS ~"
Print #1, "    Q384=+Q2   ;3. COORD. EIXO TS ~"
Print #1, "    Q333=+Q2   ;PONTO DE REFERENCIA"
Print #1, "LBL 0"

Print #1, "END PGM CentraGuias MM "
Close #1
MsgBox "Ok, CentraGuias"
'
End Sub
