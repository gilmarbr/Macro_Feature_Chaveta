Attribute VB_Name = "Inserir"
'==============================================================================
' Rasgo de chaveta DIN 6885 - ponto de entrada (botão da barra de ferramentas)
'
' Abre a janela (PropertyManager Page). Se já houver uma face cilíndrica e/ou
' um plano selecionados, eles entram preenchidos. No OK, InserirFeature cria a
' Macro Feature.
'==============================================================================
Option Explicit

' True = embute o código no arquivo da peça (a feature funciona em qualquer
' máquina, mas alterações no código só valem para features inseridas depois).
' Deixe False enquanto estiver desenvolvendo.
Private Const EMBUTIR_MACRO As Boolean = False

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' Mantém a página viva depois que main() termina
Public gPaginaInsercao As PaginaRasgo
Private gCaminhoMacro As String

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then MsgBox "Abra uma peça.", vbExclamation, "Rasgo de chaveta": Exit Sub
    If swModel.GetType <> swDocPART Then
        MsgBox "O rasgo de chaveta funciona em peças (.sldprt).", vbExclamation, "Rasgo de chaveta"
        Exit Sub
    End If

    gCaminhoMacro = swApp.GetCurrentMacroPathName
    Set gPaginaInsercao = New PaginaRasgo
    gPaginaInsercao.AbrirInsercao swApp, swModel
End Sub

' Cria a Macro Feature com as referências e parâmetros escolhidos na janela
Public Sub InserirFeature(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, _
                          face As Object, plano As Object, p As ParamsRasgo)
    Dim swFeat As SldWorks.Feature
    Dim swFace As SldWorks.Face2
    Dim g As GeoRasgo, msg As String
    Dim corpos(0) As SldWorks.Body2
    Dim metodos(8) As String, nomes() As String, tipos() As Long, valores() As String
    Dim vMetodos As Variant, vNomes As Variant, vTipos As Variant, vValores As Variant
    Dim vCorpos As Variant, vVazio As Variant, vIcones As Variant
    Dim i As Long, opcoes As Long

    ' Valida antes, para não criar uma feature já com erro
    msg = CalcularDeObjetos(face, plano, p, g)
    If msg <> "" Then MsgBox msg, vbExclamation, "Rasgo de chaveta": Exit Sub

    ' A seleção atual vira a referência da feature (face = marca 1, plano = marca 2)
    swModel.ClearSelection2 True
    SelecionarObjeto swModel, face, 1
    SelecionarObjeto swModel, plano, 2

    ' A API exige arrays tipados (String / Long), não Variant
    If gCaminhoMacro = "" Then gCaminhoMacro = swApp.GetCurrentMacroPathName
    For i = 0 To 2
        metodos(i * 3) = gCaminhoMacro
        metodos(i * 3 + 1) = "RasgoChaveta"
    Next
    metodos(2) = "swmRebuild": metodos(5) = "swmEditDefinition": metodos(8) = "swmSecurity"
    ParametrosTipados p, nomes, tipos, valores
    Set swFace = face
    Set corpos(0) = swFace.GetBody

    vMetodos = metodos: vNomes = nomes: vTipos = tipos: vValores = valores
    vCorpos = corpos
    vIcones = IconesArvore(gCaminhoMacro)

    opcoes = swMacroFeatureByDefault
    If EMBUTIR_MACRO Then opcoes = opcoes Or swMacroFeatureEmbedMacroFile

    Set swFeat = swModel.FeatureManager.InsertMacroFeature3("RasgoChaveta", "", vMetodos, _
                    vNomes, vTipos, vValores, vVazio, vVazio, vCorpos, vIcones, opcoes)

    If swFeat Is Nothing Then
        MsgBox "O SolidWorks não conseguiu inserir a feature.", vbCritical, "Rasgo de chaveta"
        Exit Sub
    End If

    Renomear swFeat, g.Designacao
    swModel.ClearSelection2 True
End Sub

' Ícones da árvore (9 bitmaps: regular/suprimido/realçado em 20, 32 e 40 px), na pasta da macro.
' Sem os arquivos, devolve Empty e o SolidWorks usa o ícone padrão.
Private Function IconesArvore(caminhoMacro As String) As Variant
    Dim pasta As String, tam As Variant, nomes(8) As String, i As Long, k As Long
    pasta = Left(caminhoMacro, InStrRev(caminhoMacro, "\"))
    For Each tam In Array("20", "32", "40")
        nomes(k) = pasta & "RasgoArvore_" & tam & ".bmp"
        nomes(k + 1) = pasta & "RasgoArvore_" & tam & "_suprimido.bmp"
        nomes(k + 2) = pasta & "RasgoArvore_" & tam & "_realce.bmp"
        k = k + 3
    Next
    For i = 0 To 8
        If Dir(nomes(i)) = "" Then Exit Function
    Next
    IconesArvore = nomes
End Function
