Attribute VB_Name = "RasgoChaveta"
'==============================================================================
' Macro Feature: Rasgo de chaveta DIN 6885
'
' O SolidWorks chama swmRebuild a cada reconstrução e swmEditDefinition ao
' editar a feature na árvore. Unidades da API: metros e radianos.
' Os parâmetros ficam guardados na feature em mm e graus.
'==============================================================================
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const MM As Double = 0.001
Private Const FOLGA As Double = 0.001          ' a ferramenta começa 1 mm fora do material
Private Const TOL As Double = 0.0000001
Private Const DEPURAR As Boolean = False      ' True = grava %TEMP%\RasgoChaveta.log

' Janela de edição (PropertyManager Page). A variável global mantém a página
' viva depois que swmEditDefinition retorna.
Public gPaginaEdicao As PaginaRasgo

Public Type ParamsRasgo
    Modo As Long        ' 0 = automático (DIN 6885), 1 = manual
    B As Double         ' largura manual (mm)
    T As Double         ' profundidade manual (mm)
    L As Double         ' comprimento (mm); 0 = automático
    Posicao As Long     ' 0 = a partir do início, 1 = centralizado, 2 = a partir do fim
    Offset As Double    ' distância da extremidade (mm)
    Angulo As Double    ' rotação em torno do eixo (graus)
    Tipo As Long        ' 0 = detectar, 1 = eixo, 2 = cubo
End Type

Public Type GeoRasgo
    Origem() As Double  ' ponto no eixo do cilindro
    Eixo() As Double    ' direção do eixo (unitária)
    U() As Double       ' direção radial em que o rasgo fica
    V() As Double       ' direção da largura do rasgo
    R As Double         ' raio do cilindro
    SMin As Double      ' extensão da face ao longo do eixo
    SMax As Double
    EhCubo As Boolean
    B As Double         ' largura (m)
    H As Double         ' altura da chaveta (m), só para a designação
    T As Double         ' profundidade (m)
    L As Double         ' comprimento (m)
    SIni As Double      ' início do rasgo ao longo do eixo (m)
    Designacao As String
End Type

'==============================================================================
' Callbacks da Macro Feature
'==============================================================================
Public Function swmRebuild(app As Variant, part As Variant, feature As Variant) As Variant
    Dim swApp As SldWorks.SldWorks
    Dim swFeat As SldWorks.feature
    Dim swData As SldWorks.MacroFeatureData
    Dim swFace As SldWorks.Face2
    Dim swBody As SldWorks.Body2
    Dim swTool As SldWorks.Body2
    Dim p As ParamsRasgo, g As GeoRasgo
    Dim refN() As Double, temRef As Boolean, msg As String
    Dim vObjs As Variant, vTypes As Variant, vMarks As Variant, vViews As Variant, vXf As Variant
    Dim vEdit As Variant, vRes As Variant, errCode As Long

    On Error GoTo Falha
    Log "rebuild: inicio"
    Set swApp = app
    Set swFeat = feature
    Set swData = swFeat.GetDefinition

    LerParametros swData, p
    swData.GetSelections3 vObjs, vTypes, vMarks, vViews, vXf
    msg = AnalisarSelecao(vObjs, vTypes, swFace, refN, temRef)
    If msg = "" Then msg = CalcularRasgo(swFace, refN, temRef, p, g)
    Log "rebuild: calc msg=[" & msg & "]"
    If msg <> "" Then swmRebuild = msg: Exit Function

    Set swTool = ConstruirFerramenta(swApp, g)
    Log "rebuild: tool=" & Not (swTool Is Nothing)
    If swTool Is Nothing Then swmRebuild = "Falha ao gerar o perfil do rasgo.": Exit Function

    vEdit = swData.EditBodies
    Log "rebuild: editBodies isArray=" & IsArray(vEdit)
    If IsEmpty(vEdit) Then swmRebuild = "Corpo a ser cortado não encontrado.": Exit Function
    Set swBody = vEdit(0)
    Set swBody = swBody.Copy

    vRes = swBody.Operations2(SWBODYCUT, swTool, errCode)
    Log "rebuild: cut isArray=" & IsArray(vRes) & " err=" & errCode
    If IsEmpty(vRes) Then swmRebuild = "O corte falhou (código " & errCode & ").": Exit Function

    AtribuirIds swData, vRes
    Log "rebuild: ok"
    If UBound(vRes) = 0 Then
        Set swmRebuild = vRes(0)
    Else
        swData.EnableMultiBodyConsume = True
        swmRebuild = vRes
    End If
    Exit Function

Falha:
    Log "rebuild: ERRO " & Err.Number & " " & Err.Description
    swmRebuild = "Rasgo de chaveta: " & Err.Description
End Function

' Diagnóstico: grava em %TEMP%\RasgoChaveta.log quando DEPURAR = True
Public Sub Log(ByVal s As String)
    Dim f As Integer
    If Not DEPURAR Then Exit Sub
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\RasgoChaveta.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & "  " & s
    Close #f
End Sub

Public Function swmEditDefinition(app As Variant, part As Variant, feature As Variant) As Variant
    Dim swModel As SldWorks.ModelDoc2
    Dim swFeat As SldWorks.feature
    Dim swData As SldWorks.MacroFeatureData

    swmEditDefinition = True
    Set swModel = part
    Set swFeat = feature
    Set swData = swFeat.GetDefinition

    If Not swData.AccessSelections(swModel, Nothing) Then
        MsgBox "Não foi possível editar a feature.", vbExclamation, "Rasgo de chaveta"
        Exit Function
    End If

    Set gPaginaEdicao = New PaginaRasgo
    gPaginaEdicao.AbrirEdicao Application.SldWorks, swModel, swFeat, swData
End Function

Public Function swmSecurity(app As Variant, part As Variant, feature As Variant) As Variant
    swmSecurity = swMacroFeatureSecurityByDefault
End Function

'==============================================================================
' Parâmetros
'==============================================================================
Public Function NomesParametros() As Variant
    NomesParametros = Array("Modo", "B", "T", "L", "Posicao", "Offset", "Angulo", "Tipo")
End Function

Public Function TiposParametros() As Variant
    TiposParametros = Array(swMacroFeatureParamTypeInteger, swMacroFeatureParamTypeDouble, _
                            swMacroFeatureParamTypeDouble, swMacroFeatureParamTypeDouble, _
                            swMacroFeatureParamTypeInteger, swMacroFeatureParamTypeDouble, _
                            swMacroFeatureParamTypeDouble, swMacroFeatureParamTypeInteger)
End Function

' Valores como texto, sempre com ponto decimal (Str ignora a configuração regional)
Public Function ValoresParametros(p As ParamsRasgo) As Variant
    ValoresParametros = Array(Trim$(Str$(p.Modo)), Trim$(Str$(p.B)), Trim$(Str$(p.T)), _
                              Trim$(Str$(p.L)), Trim$(Str$(p.Posicao)), Trim$(Str$(p.Offset)), _
                              Trim$(Str$(p.Angulo)), Trim$(Str$(p.Tipo)))
End Function

' Mesmos dados acima, em arrays tipados (exigência do InsertMacroFeature3)
Public Sub ParametrosTipados(p As ParamsRasgo, nomes() As String, tipos() As Long, valores() As String)
    Dim vN As Variant, vT As Variant, vV As Variant, i As Long
    vN = NomesParametros(): vT = TiposParametros(): vV = ValoresParametros(p)
    ReDim nomes(UBound(vN)): ReDim tipos(UBound(vN)): ReDim valores(UBound(vN))
    For i = 0 To UBound(vN)
        nomes(i) = vN(i): tipos(i) = vT(i): valores(i) = vV(i)
    Next
End Sub

Public Sub ParametrosPadrao(p As ParamsRasgo)
    p.Modo = 0: p.B = 0: p.T = 0: p.L = 0
    p.Posicao = 1: p.Offset = 0: p.Angulo = 0: p.Tipo = 0
End Sub

Public Sub LerParametros(swData As SldWorks.MacroFeatureData, p As ParamsRasgo)
    Dim i As Long, d As Double
    swData.GetIntegerByName "Modo", i: p.Modo = i
    swData.GetDoubleByName "B", d: p.B = d
    swData.GetDoubleByName "T", d: p.T = d
    swData.GetDoubleByName "L", d: p.L = d
    swData.GetIntegerByName "Posicao", i: p.Posicao = i
    swData.GetDoubleByName "Offset", d: p.Offset = d
    swData.GetDoubleByName "Angulo", d: p.Angulo = d
    swData.GetIntegerByName "Tipo", i: p.Tipo = i
End Sub

Public Sub GravarParametros(swData As SldWorks.MacroFeatureData, p As ParamsRasgo)
    swData.SetIntegerByName "Modo", p.Modo
    swData.SetDoubleByName "B", p.B
    swData.SetDoubleByName "T", p.T
    swData.SetDoubleByName "L", p.L
    swData.SetIntegerByName "Posicao", p.Posicao
    swData.SetDoubleByName "Offset", p.Offset
    swData.SetDoubleByName "Angulo", p.Angulo
    swData.SetIntegerByName "Tipo", p.Tipo
End Sub

'==============================================================================
' Seleção: 1 face cilíndrica (obrigatória) + 1 plano ou face plana (opcional)
'==============================================================================
Public Function AnalisarSelecao(vObjs As Variant, vTypes As Variant, swFace As SldWorks.Face2, _
                                refN() As Double, temRef As Boolean) As String
    Dim i As Long, f As SldWorks.Face2, s As SldWorks.Surface, vN As Variant

    Set swFace = Nothing
    temRef = False
    If Not IsArray(vObjs) Then
        AnalisarSelecao = "Selecione a face cilíndrica do eixo ou do furo do cubo."
        Exit Function
    End If

    For i = LBound(vObjs) To UBound(vObjs)
        If IsObject(vObjs(i)) Then
            If Not vObjs(i) Is Nothing Then
                Select Case vTypes(i)
                    Case swSelFACES
                        Set f = vObjs(i)
                        Set s = f.GetSurface
                        If s.IsCylinder Then
                            If swFace Is Nothing Then Set swFace = f
                        ElseIf s.IsPlane And Not temRef Then
                            vN = f.Normal
                            refN = Vec(vN(0), vN(1), vN(2))
                            temRef = True
                        End If
                    Case swSelDATUMPLANES
                        If Not temRef Then
                            refN = NormalDoPlano(vObjs(i))
                            temRef = True
                        End If
                End Select
            End If
        End If
    Next

    If swFace Is Nothing Then
        AnalisarSelecao = "Nenhuma face cilíndrica selecionada (ou a face de referência não existe mais)."
    End If
End Function

Private Function NormalDoPlano(obj As Variant) As Double()
    Dim swRefPlane As SldWorks.RefPlane, swFeat As SldWorks.feature, vT As Variant
    If TypeOf obj Is SldWorks.feature Then
        Set swFeat = obj
        Set swRefPlane = swFeat.GetSpecificFeature2
    Else
        Set swRefPlane = obj
    End If
    vT = swRefPlane.Transform.ArrayData
    NormalDoPlano = Vec(vT(6), vT(7), vT(8))     ' eixo Z local = normal do plano
End Function

'==============================================================================
' Cálculo do rasgo (dimensões, orientação e posição). Retorna "" ou o erro.
'==============================================================================
Public Function CalcularRasgo(swFace As SldWorks.Face2, refN() As Double, ByVal temRef As Boolean, _
                              p As ParamsRasgo, g As GeoRasgo) As String
    Dim vC As Variant, d As Double, compr As Double, disp As Double
    Dim bb As Double, hh As Double, t1 As Double, t2 As Double
    Dim ref() As Double, u0() As Double, v0() As Double, ang As Double, furo As Boolean

    vC = swFace.GetSurface.CylinderParams
    g.Origem = Vec(vC(0), vC(1), vC(2))
    g.Eixo = VNorm(Vec(vC(3), vC(4), vC(5)))
    g.R = vC(6)
    d = Round(2 * g.R / MM, 3)

    If Not ExtensaoAxial(swFace, g) Then
        CalcularRasgo = "Não foi possível medir o comprimento da face cilíndrica."
        Exit Function
    End If
    compr = (g.SMax - g.SMin) / MM

    ' Eixo ou cubo? A normal de uma face de eixo aponta para fora do cilindro.
    Select Case p.Tipo
        Case 1: g.EhCubo = False
        Case 2: g.EhCubo = True
        Case Else: furo = FaceEhFuro(swFace, g): g.EhCubo = furo
    End Select

    ' Dimensões
    If p.Modo = 0 Then
        If Not BuscarDIN6885(d, bb, hh, t1, t2) Then
            CalcularRasgo = "Diâmetro " & Fmt(d) & " mm fora da faixa da DIN 6885 (6 a 230 mm)."
            Exit Function
        End If
        If g.EhCubo Then g.T = t2 * MM Else g.T = t1 * MM
    Else
        If p.B <= 0 Or p.T <= 0 Then
            CalcularRasgo = "Modo manual: informe largura e profundidade maiores que zero."
            Exit Function
        End If
        bb = p.B: hh = 0
        g.T = p.T * MM
    End If
    g.B = bb * MM
    g.H = hh * MM
    If g.B / 2 >= g.R Then CalcularRasgo = "Rasgo mais largo que o diâmetro.": Exit Function
    If Not g.EhCubo And g.T >= g.R Then CalcularRasgo = "Profundidade maior que o raio do eixo.": Exit Function

    ' Orientação: o rasgo fica no lado da normal do plano de referência
    ' (ou do eixo Y global), girado pelo ângulo informado.
    If temRef Then ref = refN Else ref = Vec(0, 1, 0)
    If Abs(VDot(ref, g.Eixo)) > 0.999 Then
        If temRef Then
            CalcularRasgo = "O plano de referência deve conter o eixo (não pode ser perpendicular a ele)."
            Exit Function
        End If
        ref = Vec(0, 0, 1)
    End If
    u0 = VNorm(VSub(ref, VScl(g.Eixo, VDot(ref, g.Eixo))))
    v0 = VCross(g.Eixo, u0)
    ang = p.Angulo * 4 * Atn(1) / 180
    g.U = VAdd(VScl(u0, Cos(ang)), VScl(v0, Sin(ang)))
    g.V = VCross(g.Eixo, g.U)

    ' Posição ao longo do eixo
    If g.EhCubo Then
        g.L = (g.SMax - g.SMin) + 2 * FOLGA          ' cubo: sempre passante
        g.SIni = g.SMin - FOLGA
        If p.Modo = 0 Then
            g.Designacao = "Rasgo cubo DIN 6885 " & Fmt(bb) & "x" & Fmt(t2)
        Else
            g.Designacao = "Rasgo cubo " & Fmt(bb) & "x" & Fmt(p.T)
        End If
    Else
        If p.Posicao = 1 Then disp = compr Else disp = compr - p.Offset
        If p.L > 0 Then g.L = p.L * MM Else g.L = SugerirComprimento(d, bb, disp) * MM
        If g.L <= 0 Then
            CalcularRasgo = "Nenhum comprimento padrão cabe na face (" & Fmt(disp) & " mm disponíveis)."
            Exit Function
        End If
        If g.L < g.B - TOL Then CalcularRasgo = "Comprimento menor que a largura do rasgo.": Exit Function
        If g.L > disp * MM + TOL Then
            CalcularRasgo = "Rasgo de " & Fmt(g.L / MM) & " mm não cabe (" & Fmt(disp) & " mm disponíveis)."
            Exit Function
        End If
        Select Case p.Posicao
            Case 1: g.SIni = (g.SMin + g.SMax) / 2 - g.L / 2
            Case 2: g.SIni = g.SMax - p.Offset * MM - g.L
            Case Else: g.SIni = g.SMin + p.Offset * MM
        End Select
        If p.Modo = 0 Then
            g.Designacao = "Rasgo DIN 6885 A " & Fmt(bb) & "x" & Fmt(hh) & "x" & Fmt(g.L / MM)
        Else
            g.Designacao = "Rasgo " & Fmt(bb) & "x" & Fmt(g.L / MM) & " prof " & Fmt(p.T)
        End If
    End If
End Function

' Mede a extensão da face ao longo do eixo projetando centros de arestas
' circulares e vértices.
Private Function ExtensaoAxial(swFace As SldWorks.Face2, g As GeoRasgo) As Boolean
    Dim vE As Variant, i As Long, achou As Boolean, vP As Variant
    Dim swEdge As SldWorks.Edge, swCurve As SldWorks.Curve, swVx As SldWorks.Vertex

    vE = swFace.GetEdges
    If IsEmpty(vE) Then Exit Function
    For i = 0 To UBound(vE)
        Set swEdge = vE(i)
        Set swCurve = swEdge.GetCurve
        If swCurve.IsCircle Then
            vP = swCurve.CircleParams
            Acumular Projetar(vP(0), vP(1), vP(2), g), g, achou
        End If
        Set swVx = swEdge.GetStartVertex
        If Not swVx Is Nothing Then vP = swVx.GetPoint: Acumular Projetar(vP(0), vP(1), vP(2), g), g, achou
        Set swVx = swEdge.GetEndVertex
        If Not swVx Is Nothing Then vP = swVx.GetPoint: Acumular Projetar(vP(0), vP(1), vP(2), g), g, achou
    Next
    ExtensaoAxial = achou And (g.SMax - g.SMin) > TOL
End Function

' Eixo ou furo? As normais da tesselação apontam para fora do material:
' para longe do eixo geométrico numa face de eixo, para o eixo num furo.
Private Function FaceEhFuro(swFace As SldWorks.Face2, g As GeoRasgo) As Boolean
    Dim vT As Variant, vN As Variant, rel() As Double, rad() As Double, n() As Double
    vT = swFace.GetTessTriangles(True)
    vN = swFace.GetTessNorms
    If IsArray(vT) And IsArray(vN) Then
        rel = VSub(Vec(vT(0), vT(1), vT(2)), g.Origem)
        rad = VSub(rel, VScl(g.Eixo, VDot(rel, g.Eixo)))
        n = Vec(vN(0), vN(1), vN(2))
        FaceEhFuro = (VDot(rad, n) < 0)
        Log "furo? tess radial.n=" & Format$(VDot(rad, n) * 1000, "0.000")
    Else
        FaceEhFuro = swFace.FaceInSurfaceSense      ' reserva (observado: eixo = False)
        Log "furo? sem tesselacao, sense=" & swFace.FaceInSurfaceSense
    End If
End Function

Private Function Projetar(ByVal x As Double, ByVal y As Double, ByVal z As Double, g As GeoRasgo) As Double
    Projetar = VDot(VSub(Vec(x, y, z), g.Origem), g.Eixo)
End Function

Private Sub Acumular(ByVal s As Double, g As GeoRasgo, achou As Boolean)
    If Not achou Then
        g.SMin = s: g.SMax = s: achou = True
    Else
        If s < g.SMin Then g.SMin = s
        If s > g.SMax Then g.SMax = s
    End If
End Sub

'==============================================================================
' Corpo-ferramenta: perfil desenhado num plano tangente e extrudado para dentro
'   Eixo: oblongo (forma A, pontas arredondadas), entra no eixo.
'   Cubo: retângulo passante, cresce para fora do furo.
'==============================================================================
Public Function ConstruirFerramenta(swApp As SldWorks.SldWorks, g As GeoRasgo) As SldWorks.Body2
    Dim swMod As SldWorks.Modeler, swMath As SldWorks.MathUtility
    Dim swSurf As SldWorks.Surface, swSheet As SldWorks.Body2, swDir As SldWorks.MathVector
    Dim crv() As SldWorks.Curve, vCurvas As Variant
    Dim dirE() As Double, eixoArco() As Double, raiz() As Double
    Dim rp As Double, compr As Double, hw As Double, s1 As Double, s2 As Double

    Set swMod = swApp.GetModeler
    Set swMath = swApp.GetMathUtility
    hw = g.B / 2

    If g.EhCubo Then
        ' plano dentro do furo, abaixo da corda de largura b, para não sobrar lasca
        rp = Sqr(g.R * g.R - hw * hw) - FOLGA
        dirE = g.U
        compr = g.R + g.T - rp
    Else
        rp = g.R + FOLGA
        dirE = VScl(g.U, -1)
        compr = FOLGA + g.T
    End If

    If g.EhCubo Then
        s1 = g.SIni: s2 = g.SIni + g.L
        ReDim crv(3)
        Set crv(0) = Linha(swMod, Pt(g, s1, -hw, rp), Pt(g, s2, -hw, rp))
        Set crv(1) = Linha(swMod, Pt(g, s2, -hw, rp), Pt(g, s2, hw, rp))
        Set crv(2) = Linha(swMod, Pt(g, s2, hw, rp), Pt(g, s1, hw, rp))
        Set crv(3) = Linha(swMod, Pt(g, s1, hw, rp), Pt(g, s1, -hw, rp))
    Else
        s1 = g.SIni + hw: s2 = g.SIni + g.L - hw
        eixoArco = VScl(g.U, -1)     ' sentido anti-horário que leva -V -> +Eixo -> +V
        If s2 - s1 < TOL Then
            ReDim crv(1)             ' L = b: furo redondo (dois semicírculos)
            Set crv(0) = Arco(swMod, Pt(g, s1, 0, rp), eixoArco, hw, Pt(g, s1, -hw, rp), Pt(g, s1, hw, rp))
            Set crv(1) = Arco(swMod, Pt(g, s1, 0, rp), eixoArco, hw, Pt(g, s1, hw, rp), Pt(g, s1, -hw, rp))
        Else
            ReDim crv(3)
            Set crv(0) = Linha(swMod, Pt(g, s1, -hw, rp), Pt(g, s2, -hw, rp))
            Set crv(1) = Arco(swMod, Pt(g, s2, 0, rp), eixoArco, hw, Pt(g, s2, -hw, rp), Pt(g, s2, hw, rp))
            Set crv(2) = Linha(swMod, Pt(g, s2, hw, rp), Pt(g, s1, hw, rp))
            Set crv(3) = Arco(swMod, Pt(g, s1, 0, rp), eixoArco, hw, Pt(g, s1, hw, rp), Pt(g, s1, -hw, rp))
        End If
    End If
    Dim k As Long
    For k = 0 To UBound(crv)
        Log "tool: curva " & k & " ok=" & Not (crv(k) Is Nothing)
    Next
    vCurvas = crv

    raiz = Pt(g, g.SIni + g.L / 2, 0, rp)
    Set swSurf = swMod.CreatePlanarSurface2(AV(raiz), AV(dirE), AV(g.Eixo))
    Log "tool: surf=" & Not (swSurf Is Nothing)
    If swSurf Is Nothing Then Exit Function
    Set swSheet = swSurf.CreateTrimmedSheet4(vCurvas, True)
    Log "tool: sheet=" & Not (swSheet Is Nothing)
    If swSheet Is Nothing Then Exit Function

    Set swDir = swMath.CreateVector(AV(dirE))
    Set ConstruirFerramenta = swMod.CreateExtrudedBody(swSheet, swDir, compr)
    Log "tool: extrude=" & Not (ConstruirFerramenta Is Nothing)
End Function

Private Function Linha(swMod As SldWorks.Modeler, a() As Double, b() As Double) As SldWorks.Curve
    Dim c As SldWorks.Curve
    Set c = swMod.CreateLine(AV(a), AV(VSub(b, a)))
    Set Linha = c.CreateTrimmedCurve2(a(0), a(1), a(2), b(0), b(1), b(2))
End Function

Private Function Arco(swMod As SldWorks.Modeler, centro() As Double, eixo() As Double, ByVal raio As Double, _
                      a() As Double, b() As Double) As SldWorks.Curve
    Dim c As SldWorks.Curve
    Set c = swMod.CreateArc(AV(centro), AV(eixo), raio, AV(a), AV(b))
    Set Arco = c.CreateTrimmedCurve2(a(0), a(1), a(2), b(0), b(1), b(2))
End Function

' IDs estáveis para faces e arestas novas (evita referências perdidas em
' cotas e relações de montagem feitas sobre o rasgo)
Private Sub AtribuirIds(swData As SldWorks.MacroFeatureData, vRes As Variant)
    Dim i As Long, j As Long, nF As Long, nE As Long
    Dim swBody As SldWorks.Body2, vF As Variant, vE As Variant
    On Error Resume Next
    For i = 0 To UBound(vRes)
        Set swBody = vRes(i)
        vF = Empty: vE = Empty
        swData.GetEntitiesNeedUserId swBody, vF, vE
        If IsArray(vF) Then
            For j = 0 To UBound(vF)
                swData.SetFaceUserId vF(j), nF, 0: nF = nF + 1
            Next
        End If
        If IsArray(vE) Then
            For j = 0 To UBound(vE)
                swData.SetEdgeUserId vE(j), nE, 0: nE = nE + 1
            Next
        End If
    Next
End Sub

'==============================================================================
' Helpers usados pela PropertyManager Page e pela inserção
'==============================================================================
' Calcula o rasgo a partir dos objetos selecionados (face + plano opcional)
Public Function CalcularDeObjetos(face As Object, plano As Object, p As ParamsRasgo, g As GeoRasgo) As String
    Dim vO As Variant, vT As Variant, swFace As SldWorks.Face2, refN() As Double, temRef As Boolean
    Dim msg As String
    If face Is Nothing Then
        CalcularDeObjetos = "Selecione a face cilíndrica do eixo ou do furo."
        Exit Function
    End If
    If plano Is Nothing Then
        vO = Array(face): vT = Array(CLng(swSelFACES))
    Else
        vO = Array(face, plano): vT = Array(CLng(swSelFACES), TipoSelecao(plano))
    End If
    msg = AnalisarSelecao(vO, vT, swFace, refN, temRef)
    If msg = "" Then msg = CalcularRasgo(swFace, refN, temRef, p, g)
    CalcularDeObjetos = msg
End Function

Public Function TipoSelecao(obj As Object) As Long
    If TypeOf obj Is SldWorks.Face2 Then TipoSelecao = swSelFACES Else TipoSelecao = swSelDATUMPLANES
End Function

' Seleciona (acrescentando) uma face ou um plano com a marca indicada
Public Sub SelecionarObjeto(swModel As SldWorks.ModelDoc2, obj As Object, ByVal marca As Long)
    Dim sd As SldWorks.SelectData, swEnt As SldWorks.Entity, swF As SldWorks.feature
    If obj Is Nothing Then Exit Sub
    If TypeOf obj Is SldWorks.feature Then
        Set swF = obj
        swF.Select2 True, marca
    Else
        Set sd = swModel.SelectionManager.CreateSelectData
        sd.Mark = marca
        Set swEnt = obj
        swEnt.Select4 True, sd
    End If
End Sub
'==============================================================================
' Utilidades
'==============================================================================
Public Sub Renomear(swFeat As SldWorks.feature, ByVal nome As String)
    Dim k As Long, tentativa As String
    On Error Resume Next
    tentativa = nome
    For k = 2 To 30
        swFeat.Name = tentativa
        If swFeat.Name = tentativa Then Exit Sub
        tentativa = nome & " (" & k & ")"
    Next
End Sub

Public Function Fmt(ByVal x As Double) As String
    Dim s As String
    s = Format$(Round(x, 2), "0.##")
    If Right$(s, 1) = "," Or Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    Fmt = s
End Function

Private Function Pt(g As GeoRasgo, ByVal s As Double, ByVal w As Double, ByVal rp As Double) As Double()
    Pt = VAdd(VAdd(VAdd(g.Origem, VScl(g.Eixo, s)), VScl(g.U, rp)), VScl(g.V, w))
End Function

Private Function AV(a() As Double) As Variant
    AV = a
End Function

Public Function Vec(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Double()
    Dim r(2) As Double
    r(0) = x: r(1) = y: r(2) = z
    Vec = r
End Function

Private Function VAdd(a() As Double, b() As Double) As Double()
    VAdd = Vec(a(0) + b(0), a(1) + b(1), a(2) + b(2))
End Function

Private Function VSub(a() As Double, b() As Double) As Double()
    VSub = Vec(a(0) - b(0), a(1) - b(1), a(2) - b(2))
End Function

Private Function VScl(a() As Double, ByVal k As Double) As Double()
    VScl = Vec(a(0) * k, a(1) * k, a(2) * k)
End Function

Private Function VDot(a() As Double, b() As Double) As Double
    VDot = a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
End Function

Private Function VCross(a() As Double, b() As Double) As Double()
    VCross = Vec(a(1) * b(2) - a(2) * b(1), a(2) * b(0) - a(0) * b(2), a(0) * b(1) - a(1) * b(0))
End Function

Private Function VNorm(a() As Double) As Double()
    Dim n As Double
    n = Sqr(VDot(a, a))
    VNorm = VScl(a, 1 / n)
End Function
