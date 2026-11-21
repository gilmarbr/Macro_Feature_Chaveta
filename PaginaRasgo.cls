VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PaginaRasgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PropertyManager Page do Rasgo de chaveta DIN 6885
'
' Usada na inserção (Inserir.main) e na edição (swmEditDefinition).
' Requer a referência "SOLIDWORKS <versão> exposed type libraries for add-in use".
'==============================================================================
Option Explicit

Implements PropertyManagerPage2Handler9

Private Const PI As Double = 3.14159265358979

' IDs dos controles
Private Const ID_GRP_REF As Long = 1
Private Const ID_GRP_TIPO As Long = 2
Private Const ID_GRP_DIM As Long = 3
Private Const ID_GRP_POS As Long = 4
Private Const ID_SEL_FACE As Long = 10
Private Const ID_SEL_PLANO As Long = 11
Private Const ID_OPT_DETECTAR As Long = 20
Private Const ID_OPT_EIXO As Long = 21
Private Const ID_OPT_CUBO As Long = 22
Private Const ID_OPT_DIN As Long = 30
Private Const ID_OPT_MANUAL As Long = 31
Private Const ID_LBL_B As Long = 32
Private Const ID_NUM_B As Long = 33
Private Const ID_LBL_T As Long = 34
Private Const ID_NUM_T As Long = 35
Private Const ID_LBL_INFO As Long = 36
Private Const ID_LBL_POS As Long = 40
Private Const ID_CBO_POS As Long = 41
Private Const ID_LBL_DIST As Long = 42
Private Const ID_NUM_DIST As Long = 43
Private Const ID_CHK_AUTO As Long = 44
Private Const ID_LBL_L As Long = 45
Private Const ID_NUM_L As Long = 46
Private Const ID_LBL_ANG As Long = 47
Private Const ID_NUM_ANG As Long = 48
Private Const ID_BTN_INV As Long = 49

Private Const MARCA_FACE As Long = 1
Private Const MARCA_PLANO As Long = 2

' Estado
Private swApp As SldWorks.SldWorks
Private swModel As SldWorks.ModelDoc2
Private swPage As SldWorks.PropertyManagerPage2
Private swFeat As SldWorks.Feature
Private swData As SldWorks.MacroFeatureData
Private mEdicao As Boolean
Private mP As ParamsRasgo
Private mFace As Object
Private mPlano As Object
Private mOk As Boolean
Private mPronto As Boolean
Private mPreview As SldWorks.Body2
Private mAberta As Boolean

' Controles (os grupos também precisam ficar referenciados enquanto a página existir)
Private grpRef As SldWorks.PropertyManagerPageGroup
Private grpTipo As SldWorks.PropertyManagerPageGroup
Private grpDim As SldWorks.PropertyManagerPageGroup
Private grpPos As SldWorks.PropertyManagerPageGroup
Private selFace As SldWorks.PropertyManagerPageSelectionbox
Private selPlano As SldWorks.PropertyManagerPageSelectionbox
Private optDetectar As SldWorks.PropertyManagerPageOption
Private optEixo As SldWorks.PropertyManagerPageOption
Private optCubo As SldWorks.PropertyManagerPageOption
Private optDin As SldWorks.PropertyManagerPageOption
Private optManual As SldWorks.PropertyManagerPageOption
Private lblB As SldWorks.PropertyManagerPageLabel
Private numB As SldWorks.PropertyManagerPageNumberbox
Private lblT As SldWorks.PropertyManagerPageLabel
Private numT As SldWorks.PropertyManagerPageNumberbox
Private lblInfo As SldWorks.PropertyManagerPageLabel
Private lblPos As SldWorks.PropertyManagerPageLabel
Private cboPos As SldWorks.PropertyManagerPageCombobox
Private lblDist As SldWorks.PropertyManagerPageLabel
Private numDist As SldWorks.PropertyManagerPageNumberbox
Private chkAuto As SldWorks.PropertyManagerPageCheckbox
Private lblL As SldWorks.PropertyManagerPageLabel
Private numL As SldWorks.PropertyManagerPageNumberbox
Private lblAng As SldWorks.PropertyManagerPageLabel
Private numAng As SldWorks.PropertyManagerPageNumberbox
Private btnInv As SldWorks.PropertyManagerPageButton

'==============================================================================
' Abertura
'==============================================================================
' True enquanto a janela estiver aberta (quem abriu espera com DoEvents)
Public Property Get Aberta() As Boolean
    Aberta = mAberta
End Property

Public Sub AbrirInsercao(app As SldWorks.SldWorks, model As SldWorks.ModelDoc2)
    Dim swSel As SldWorks.SelectionMgr, i As Long
    Set swApp = app
    Set swModel = model
    mEdicao = False
    ParametrosPadrao mP

    ' Aproveita o que já estava selecionado (face cilíndrica / plano)
    Set swSel = swModel.SelectionManager
    For i = 1 To swSel.GetSelectedObjectCount2(-1)
        Classificar swSel.GetSelectedObject6(i, -1), swSel.GetSelectedObjectType3(i, -1)
    Next
    swModel.ClearSelection2 True
    Mostrar
End Sub

Public Sub AbrirEdicao(app As SldWorks.SldWorks, model As SldWorks.ModelDoc2, _
                       feat As SldWorks.Feature, data As SldWorks.MacroFeatureData)
    Dim vObjs As Variant, vTypes As Variant, vMarks As Variant, vViews As Variant, vXf As Variant, i As Long
    Set swApp = app
    Set swModel = model
    Set swFeat = feat
    Set swData = data
    mEdicao = True
    LerParametros swData, mP

    swData.GetSelections3 vObjs, vTypes, vMarks, vViews, vXf
    If IsArray(vObjs) Then
        For i = LBound(vObjs) To UBound(vObjs)
            If IsObject(vObjs(i)) Then Classificar vObjs(i), vTypes(i)
        Next
    End If
    swModel.ClearSelection2 True
    Mostrar
End Sub

Private Sub Classificar(ByVal obj As Object, ByVal tipo As Long)
    Dim f As SldWorks.Face2
    If obj Is Nothing Then Exit Sub
    If tipo = swSelFACES Then
        Set f = obj
        If f.GetSurface.IsCylinder Then
            If mFace Is Nothing Then Set mFace = obj
        ElseIf f.GetSurface.IsPlane Then
            If mPlano Is Nothing Then Set mPlano = obj
        End If
    ElseIf tipo = swSelDATUMPLANES Then
        If mPlano Is Nothing Then Set mPlano = obj
    End If
End Sub

Private Sub Mostrar()
    Log "pmp: criando pagina (edicao=" & mEdicao & ")"
    If Not CriarPagina() Then
        mAberta = False
        MsgBox "Não foi possível criar a janela do rasgo de chaveta.", vbCritical, "Rasgo de chaveta"
        If mEdicao Then swData.ReleaseSelectionAccess
        Exit Sub
    End If
    Log "pmp: pagina criada"
    PreencherControles
    ' Pré-seleção com as marcas das caixas: o SolidWorks distribui ao exibir
    swModel.ClearSelection2 True
    SelecionarObjeto swModel, mFace, MARCA_FACE
    SelecionarObjeto swModel, mPlano, MARCA_PLANO
    Log "pmp: pre-selecao feita"
    mAberta = True
    swPage.Show
    Log "pmp: Show ok"
    ' Primeira atualização só depois de a página estar montada
    mPronto = True
    Atualizar
End Sub

'==============================================================================
' Montagem da página
'==============================================================================
Private Function CriarPagina() As Boolean
    Dim errs As Long, opc As Long, grpOpc As Long, ctl As Long
    Dim grp As SldWorks.PropertyManagerPageGroup   ' atalho local; o grupo fica guardado em grpXxx
    Dim filtros() As Long, itens() As String, v As Variant

    opc = swPropertyManagerOptions_OkayButton + swPropertyManagerOptions_CancelButton
    Set swPage = swApp.CreatePropertyManagerPage("Rasgo de chaveta DIN 6885", opc, Me, errs)
    If swPage Is Nothing Or errs <> swPropertyManagerPage_Okay Then Exit Function

    grpOpc = swGroupBoxOptions_Visible + swGroupBoxOptions_Expanded
    ctl = swControlOptions_Visible + swControlOptions_Enabled

    ' --- Referencias -------------------------------------------------------
    Set grpRef = swPage.AddGroupBox(ID_GRP_REF, "Referências", grpOpc): Set grp = grpRef
    Set selFace = grp.AddControl(ID_SEL_FACE, CInt(swControlType_Selectionbox), "Face cilíndrica", _
                  swControlAlign_Indent, ctl, "Face cilíndrica do eixo ou do furo do cubo")
    ReDim filtros(0): filtros(0) = swSelFACES
    v = filtros: selFace.SetSelectionFilters v
    selFace.SingleEntityOnly = True
    selFace.Mark = MARCA_FACE
    selFace.Height = 20
    Etiqueta selFace, swBitmapLabel_SelectFace

    Set selPlano = grp.AddControl(ID_SEL_PLANO, CInt(swControlType_Selectionbox), "Plano de orientação", _
                   swControlAlign_Indent, ctl, "Opcional: plano ou face plana que contenha o eixo. " & _
                   "O rasgo fica do lado da normal.")
    ReDim filtros(1): filtros(0) = swSelDATUMPLANES: filtros(1) = swSelFACES
    v = filtros: selPlano.SetSelectionFilters v
    selPlano.SingleEntityOnly = True
    selPlano.Mark = MARCA_PLANO
    selPlano.Height = 20
    Etiqueta selPlano, swBitmapLabel_SelectFaceSurface

    ' --- Aplicar em --------------------------------------------------------
    Set grpTipo = swPage.AddGroupBox(ID_GRP_TIPO, "Aplicar em", grpOpc): Set grp = grpTipo
    Set optDetectar = grp.AddControl(ID_OPT_DETECTAR, CInt(swControlType_Option), "Detectar automaticamente", _
                      swControlAlign_Indent, ctl, "Decide pela face: externa = eixo, interna = cubo")
    Set optEixo = grp.AddControl(ID_OPT_EIXO, CInt(swControlType_Option), "Eixo (rasgo externo, forma A)", _
                  swControlAlign_Indent, ctl, "Rasgo com pontas arredondadas, profundidade t1")
    Set optCubo = grp.AddControl(ID_OPT_CUBO, CInt(swControlType_Option), "Cubo (rasgo no furo, passante)", _
                  swControlAlign_Indent, ctl, "Rasgo retangular passante, profundidade t2")

    ' --- Dimensões ---------------------------------------------------------
    Set grpDim = swPage.AddGroupBox(ID_GRP_DIM, "Dimensões", grpOpc): Set grp = grpDim
    Set optDin = grp.AddControl(ID_OPT_DIN, CInt(swControlType_Option), "Automático pela DIN 6885", _
                 swControlAlign_Indent, ctl, "Largura e profundidade pela tabela da norma, a partir do diâmetro")
    Set optManual = grp.AddControl(ID_OPT_MANUAL, CInt(swControlType_Option), "Manual", _
                    swControlAlign_Indent, ctl, "Informe largura e profundidade")
    Set lblB = grp.AddControl(ID_LBL_B, CInt(swControlType_Label), "Largura b:", CInt(swControlAlign_Indent), ctl, "")
    Set numB = grp.AddControl(ID_NUM_B, CInt(swControlType_Numberbox), "Largura", CInt(swControlAlign_Indent), ctl, "Largura do rasgo")
    numB.SetRange2 swNumberBox_Length, 0.0005, 0.2, True, 0.001, 0.005, 0.0001
    Etiqueta numB, swBitmapLabel_Width
    Set lblT = grp.AddControl(ID_LBL_T, CInt(swControlType_Label), "Profundidade t:", CInt(swControlAlign_Indent), ctl, "")
    Set numT = grp.AddControl(ID_NUM_T, CInt(swControlType_Numberbox), "Profundidade", CInt(swControlAlign_Indent), ctl, _
               "Profundidade a partir da superfície, na linha de centro")
    numT.SetRange2 swNumberBox_Length, 0.0002, 0.1, True, 0.0005, 0.002, 0.0001
    Etiqueta numT, swBitmapLabel_Depth
    Set lblInfo = grp.AddControl(ID_LBL_INFO, CInt(swControlType_Label), " ", CInt(swControlAlign_LeftEdge), ctl, "")

    ' --- Posição -----------------------------------------------------------
    Set grpPos = swPage.AddGroupBox(ID_GRP_POS, "Posição", grpOpc): Set grp = grpPos
    Set lblPos = grp.AddControl(ID_LBL_POS, CInt(swControlType_Label), "Referência ao longo do eixo:", CInt(swControlAlign_Indent), ctl, "")
    Set cboPos = grp.AddControl(ID_CBO_POS, CInt(swControlType_Combobox), "Referência", CInt(swControlAlign_Indent), ctl, "")
    ReDim itens(2): itens(0) = "A partir do início da face": itens(1) = "Centralizado": itens(2) = "A partir do fim da face"
    v = itens: cboPos.AddItems v
    Set lblDist = grp.AddControl(ID_LBL_DIST, CInt(swControlType_Label), "Distância da extremidade:", CInt(swControlAlign_Indent), ctl, "")
    Set numDist = grp.AddControl(ID_NUM_DIST, CInt(swControlType_Numberbox), "Distância", CInt(swControlAlign_Indent), ctl, "")
    numDist.SetRange2 swNumberBox_Length, 0, 2, True, 0.001, 0.005, 0.0001
    Etiqueta numDist, swBitmapLabel_LinearDistance
    Set chkAuto = grp.AddControl(ID_CHK_AUTO, CInt(swControlType_Checkbox), "Comprimento padrão automático (~1,5 x d)", _
                  swControlAlign_Indent, ctl, "Escolhe o comprimento padronizado mais próximo de 1,5 x d que caiba na face")
    Set lblL = grp.AddControl(ID_LBL_L, CInt(swControlType_Label), "Comprimento L:", CInt(swControlAlign_Indent), ctl, "")
    Set numL = grp.AddControl(ID_NUM_L, CInt(swControlType_Numberbox), "Comprimento", CInt(swControlAlign_Indent), ctl, "")
    numL.SetRange2 swNumberBox_Length, 0.001, 2, True, 0.001, 0.005, 0.0001
    Etiqueta numL, swBitmapLabel_LinearDistance1
    Set lblAng = grp.AddControl(ID_LBL_ANG, CInt(swControlType_Label), "Ângulo em torno do eixo:", CInt(swControlAlign_Indent), ctl, "")
    Set numAng = grp.AddControl(ID_NUM_ANG, CInt(swControlType_Numberbox), "Ângulo", CInt(swControlAlign_Indent), ctl, "")
    numAng.SetRange2 swNumberBox_Angle, -2 * PI, 2 * PI, True, PI / 12, PI / 4, PI / 180
    Etiqueta numAng, swBitmapLabel_AngularDistance
    Set btnInv = grp.AddControl(ID_BTN_INV, CInt(swControlType_Button), "Inverter lado (180°)", _
                 swControlAlign_Indent, ctl, "Gira o rasgo para o lado oposto do eixo")

    Log "pmp: tipos grp=" & TypeName(grp) & " sel=" & TypeName(selFace) & " num=" & TypeName(numB) & " lbl=" & TypeName(lblB)
    CriarPagina = True
End Function

Private Sub Etiqueta(ctl As Object, ByVal tipo As Long)
    Exit Sub   ' desativado: a conversão para PropertyManagerPageControl é instável no VBA
    Dim c As SldWorks.PropertyManagerPageControl
    On Error Resume Next
    Set c = ctl
    c.SetStandardPictureLabel tipo
End Sub

' Desativado: converter um controle (em especial um rótulo) para
' PropertyManagerPageControl e mudar Enabled durante os eventos derruba o
' SolidWorks 2024 a partir do VBA. Os campos ficam sempre editáveis.
Private Sub Habilitar(ctl As Object, ByVal ligado As Boolean)
    Exit Sub
    Dim c As SldWorks.PropertyManagerPageControl
    On Error Resume Next
    Set c = ctl
    c.Enabled = ligado
End Sub

'==============================================================================
' Controles <-> parâmetros
'==============================================================================
Private Sub PreencherControles()
    mPronto = False
    Select Case mP.Tipo
        Case 1: optEixo.Checked = True
        Case 2: optCubo.Checked = True
        Case Else: optDetectar.Checked = True
    End Select
    If mP.Modo = 1 Then optManual.Checked = True Else optDin.Checked = True
    If mP.B > 0 Then numB.Value = mP.B * 0.001 Else numB.Value = 0.008
    If mP.T > 0 Then numT.Value = mP.T * 0.001 Else numT.Value = 0.004
    cboPos.CurrentSelection = mP.Posicao
    numDist.Value = mP.Offset * 0.001
    chkAuto.Checked = (mP.L <= 0)
    If mP.L > 0 Then numL.Value = mP.L * 0.001 Else numL.Value = 0.04
    numAng.Value = mP.Angulo * PI / 180
    mPronto = True
End Sub

Private Sub LerControles()
    If optEixo.Checked Then
        mP.Tipo = 1
    ElseIf optCubo.Checked Then
        mP.Tipo = 2
    Else
        mP.Tipo = 0
    End If
    If optManual.Checked Then mP.Modo = 1 Else mP.Modo = 0
    mP.B = Round(numB.Value * 1000, 4)
    mP.T = Round(numT.Value * 1000, 4)
    mP.Posicao = cboPos.CurrentSelection
    If mP.Posicao < 0 Or mP.Posicao > 2 Then mP.Posicao = 1
    mP.Offset = Round(numDist.Value * 1000, 4)
    If chkAuto.Checked Then mP.L = 0 Else mP.L = Round(numL.Value * 1000, 4)
    mP.Angulo = Round(numAng.Value * 180 / PI, 4)
End Sub

'==============================================================================
' Atualização (habilitação + pré-visualização)
'==============================================================================
Private Sub Atualizar()
    Dim g As GeoRasgo, msg As String, face As Object, plano As Object
    If Not mPronto Then Exit Sub
    On Error GoTo Falha

    Log "upd: inicio"
    LerControles
    Log "upd: controles lidos"
    EsconderPreview

    Habilitar lblB, optManual.Checked: Habilitar numB, optManual.Checked
    Habilitar lblT, optManual.Checked: Habilitar numT, optManual.Checked
    Log "upd: habilitacao ok"

    Set face = ObterSelecao(MARCA_FACE)
    Set plano = ObterSelecao(MARCA_PLANO)
    Log "upd: face=" & (Not face Is Nothing) & " plano=" & (Not plano Is Nothing)
    If face Is Nothing Then
        lblInfo.Caption = "Selecione a face cilíndrica do eixo ou do furo."
        HabilitarPosicao True
        Exit Sub
    End If

    msg = CalcularDeObjetos(face, plano, mP, g)
    Log "upd: calc [" & msg & "]"
    If msg <> "" Then
        lblInfo.Caption = "! " & msg
        HabilitarPosicao Not optCubo.Checked
        Exit Sub
    End If

    HabilitarPosicao Not g.EhCubo
    If g.EhCubo Then
        lblInfo.Caption = g.Designacao & "  (passante)"
    Else
        lblInfo.Caption = g.Designacao & "  -  prof. " & Fmt(g.T * 1000) & " mm"
        If chkAuto.Checked Then
            mPronto = False
            numL.Value = g.L
            mPronto = True
        End If
    End If
    If optDin.Checked Then
        mPronto = False
        numB.Value = g.B
        numT.Value = g.T
        mPronto = True
    End If

    Log "upd: controles atualizados"
    Set mPreview = ConstruirFerramenta(swApp, g)
    Log "upd: ferramenta=" & Not (mPreview Is Nothing)
    If Not mPreview Is Nothing Then
        mPreview.Display3 swModel, RGB(255, 170, 0), swTempBodySelectOptionNone
        Log "upd: preview exibido"
        swModel.GraphicsRedraw2
    End If
    Exit Sub

Falha:
    Log "pmp: Atualizar ERRO " & Err.Description
    lblInfo.Caption = "! " & Err.Description
    mPronto = True
End Sub

Private Sub HabilitarPosicao(ByVal ligado As Boolean)
    Habilitar lblPos, ligado: Habilitar cboPos, ligado
    Habilitar lblDist, ligado And cboPos.CurrentSelection <> 1
    Habilitar numDist, ligado And cboPos.CurrentSelection <> 1
    Habilitar chkAuto, ligado
    Habilitar lblL, ligado And Not chkAuto.Checked
    Habilitar numL, ligado And Not chkAuto.Checked
End Sub

Private Sub EsconderPreview()
    On Error Resume Next
    If Not mPreview Is Nothing Then
        mPreview.Hide swModel
        Set mPreview = Nothing
        swModel.GraphicsRedraw2
    End If
End Sub

Private Function ObterSelecao(ByVal marca As Long) As Object
    Dim swSel As SldWorks.SelectionMgr
    Set swSel = swModel.SelectionManager
    If swSel.GetSelectedObjectCount2(marca) > 0 Then Set ObterSelecao = swSel.GetSelectedObject6(1, marca)
End Function

'==============================================================================
' Conclusão
'==============================================================================
Private Sub Concluir()
    Dim g As GeoRasgo, msg As String
    Dim objs() As Object, marcas() As Long, vistas() As SldWorks.View, n As Long
    Dim vO As Variant, vM As Variant, vV As Variant

    If Not mEdicao Then
        If mOk Then
            If mFace Is Nothing Then
                MsgBox "Nenhuma face cilíndrica selecionada.", vbExclamation, "Rasgo de chaveta"
            Else
                InserirFeature swApp, swModel, mFace, mPlano, mP
            End If
        End If
        Exit Sub
    End If

    If Not mOk Or mFace Is Nothing Then
        swData.ReleaseSelectionAccess
        Exit Sub
    End If

    msg = CalcularDeObjetos(mFace, mPlano, mP, g)
    If mPlano Is Nothing Then n = 0 Else n = 1
    ReDim objs(n): ReDim marcas(n): ReDim vistas(n)
    Set objs(0) = mFace: marcas(0) = MARCA_FACE
    If n = 1 Then Set objs(1) = mPlano: marcas(1) = MARCA_PLANO
    vO = objs: vM = marcas: vV = vistas
    swData.SetSelections2 vO, vM, vV
    GravarParametros swData, mP
    swFeat.ModifyDefinition swData, swModel, Nothing
    If msg = "" Then Renomear swFeat, g.Designacao
End Sub

'==============================================================================
' Eventos da PropertyManager Page
'==============================================================================
Private Sub PropertyManagerPage2Handler9_AfterActivation()
    On Error GoTo Falha
    Log "pmp: AfterActivation"
    Log "pmp: AfterActivation fim"
    Exit Sub
Falha:
    mPronto = True
    Log "pmp: AfterActivation ERRO " & Err.Number & " " & Err.Description
End Sub

Private Sub PropertyManagerPage2Handler9_OnClose(ByVal Reason As Long)
    Log "pmp: OnClose " & Reason
    On Error GoTo Falha
    mPronto = False
    EsconderPreview
    If Reason = swPropertyManagerPageClose_Okay Then
        mPronto = False
        LerControles
        Set mFace = ObterSelecao(MARCA_FACE)
        Set mPlano = ObterSelecao(MARCA_PLANO)
        mOk = True
    Else
        mOk = False
    End If
    Exit Sub
Falha:
    Log "pmp: OnClose ERRO " & Err.Description
End Sub

Private Sub PropertyManagerPage2Handler9_AfterClose()
    Log "pmp: AfterClose ok=" & mOk
    On Error GoTo Falha
    Concluir
    mAberta = False
    Exit Sub
Falha:
    mAberta = False
    Log "pmp: ERRO " & Err.Description
    MsgBox "Rasgo de chaveta: " & Err.Description, vbExclamation, "Rasgo de chaveta"
End Sub

Private Function PropertyManagerPage2Handler9_OnSubmitSelection(ByVal Id As Long, ByVal Selection As Object, _
                                                               ByVal SelType As Long, ItemText As String) As Boolean
    Dim f As SldWorks.Face2
    On Error GoTo Nao
    Select Case Id
        Case ID_SEL_FACE
            If SelType = swSelFACES Then
                Set f = Selection
                PropertyManagerPage2Handler9_OnSubmitSelection = f.GetSurface.IsCylinder
            End If
        Case ID_SEL_PLANO
            If SelType = swSelDATUMPLANES Then
                PropertyManagerPage2Handler9_OnSubmitSelection = True
            ElseIf SelType = swSelFACES Then
                Set f = Selection
                PropertyManagerPage2Handler9_OnSubmitSelection = f.GetSurface.IsPlane
            End If
        Case Else
            PropertyManagerPage2Handler9_OnSubmitSelection = True
    End Select
    Exit Function
Nao:
    PropertyManagerPage2Handler9_OnSubmitSelection = False
End Function

Private Sub PropertyManagerPage2Handler9_OnSelectionboxListChanged(ByVal Id As Long, ByVal Count As Long)
    Atualizar
End Sub

Private Sub PropertyManagerPage2Handler9_OnOptionCheck(ByVal Id As Long)
    Atualizar
End Sub

Private Sub PropertyManagerPage2Handler9_OnCheckboxCheck(ByVal Id As Long, ByVal Checked As Boolean)
    Atualizar
End Sub

Private Sub PropertyManagerPage2Handler9_OnNumberboxChanged(ByVal Id As Long, ByVal Value As Double)
    Atualizar
End Sub

Private Sub PropertyManagerPage2Handler9_OnComboboxSelectionChanged(ByVal Id As Long, ByVal Item As Long)
    Atualizar
End Sub

Private Sub PropertyManagerPage2Handler9_OnButtonPress(ByVal Id As Long)
    Dim a As Double
    If Id <> ID_BTN_INV Then Exit Sub
    a = numAng.Value + PI
    If a > PI + 0.000001 Then a = a - 2 * PI
    mPronto = False
    numAng.Value = a
    mPronto = True
    Atualizar
End Sub

' --- Eventos não usados (exigidos pela interface) ---------------------------
Private Function PropertyManagerPage2Handler9_OnActiveXControlCreated(ByVal Id As Long, ByVal Status As Boolean) As Long
    PropertyManagerPage2Handler9_OnActiveXControlCreated = 0
End Function
Private Sub PropertyManagerPage2Handler9_OnComboboxEditChanged(ByVal Id As Long, ByVal Text As String)
End Sub
Private Sub PropertyManagerPage2Handler9_OnGainedFocus(ByVal Id As Long)
End Sub
Private Sub PropertyManagerPage2Handler9_OnGroupCheck(ByVal Id As Long, ByVal Checked As Boolean)
End Sub
Private Sub PropertyManagerPage2Handler9_OnGroupExpand(ByVal Id As Long, ByVal Expanded As Boolean)
End Sub
Private Function PropertyManagerPage2Handler9_OnHelp() As Boolean
    PropertyManagerPage2Handler9_OnHelp = True
End Function
Private Function PropertyManagerPage2Handler9_OnKeystroke(ByVal Wparam As Long, ByVal Message As Long, _
                                                         ByVal Lparam As Long, ByVal Id As Long) As Boolean
    PropertyManagerPage2Handler9_OnKeystroke = False
End Function
Private Sub PropertyManagerPage2Handler9_OnListboxRMBUp(ByVal Id As Long, ByVal PosX As Long, ByVal PosY As Long)
End Sub
Private Sub PropertyManagerPage2Handler9_OnListboxSelectionChanged(ByVal Id As Long, ByVal Item As Long)
End Sub
Private Sub PropertyManagerPage2Handler9_OnLostFocus(ByVal Id As Long)
End Sub
Private Function PropertyManagerPage2Handler9_OnNextPage() As Boolean
    PropertyManagerPage2Handler9_OnNextPage = True
End Function
Private Sub PropertyManagerPage2Handler9_OnNumberBoxTrackingCompleted(ByVal Id As Long, ByVal Value As Double)
End Sub
Private Sub PropertyManagerPage2Handler9_OnPopupMenuItem(ByVal Id As Long)
End Sub
Private Sub PropertyManagerPage2Handler9_OnPopupMenuItemUpdate(ByVal Id As Long, retval As Long)
End Sub
Private Function PropertyManagerPage2Handler9_OnPreview() As Boolean
    PropertyManagerPage2Handler9_OnPreview = True
End Function
Private Function PropertyManagerPage2Handler9_OnPreviousPage() As Boolean
    PropertyManagerPage2Handler9_OnPreviousPage = True
End Function
Private Sub PropertyManagerPage2Handler9_OnRedo()
End Sub
Private Sub PropertyManagerPage2Handler9_OnSelectionboxCalloutCreated(ByVal Id As Long)
End Sub
Private Sub PropertyManagerPage2Handler9_OnSelectionboxCalloutDestroyed(ByVal Id As Long)
End Sub
Private Sub PropertyManagerPage2Handler9_OnSelectionboxFocusChanged(ByVal Id As Long)
End Sub
Private Sub PropertyManagerPage2Handler9_OnSliderPositionChanged(ByVal Id As Long, ByVal Value As Double)
End Sub
Private Sub PropertyManagerPage2Handler9_OnSliderTrackingCompleted(ByVal Id As Long, ByVal Value As Double)
End Sub
Private Function PropertyManagerPage2Handler9_OnTabClicked(ByVal Id As Long) As Boolean
    PropertyManagerPage2Handler9_OnTabClicked = True
End Function
Private Sub PropertyManagerPage2Handler9_OnTextboxChanged(ByVal Id As Long, ByVal Text As String)
End Sub
Private Sub PropertyManagerPage2Handler9_OnUndo()
End Sub
Private Sub PropertyManagerPage2Handler9_OnWhatsNew()
End Sub
Private Function PropertyManagerPage2Handler9_OnWindowFromHandleControlCreated(ByVal Id As Long, _
                                                                              ByVal Status As Boolean) As Long
    PropertyManagerPage2Handler9_OnWindowFromHandleControlCreated = 0
End Function
