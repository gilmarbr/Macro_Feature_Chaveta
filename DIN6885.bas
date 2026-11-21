Attribute VB_Name = "DIN6885"
'==============================================================================
' Tabela DIN 6885-1 - chaveta paralela, forma A. Valores em mm.
'   d  = di?metro do eixo
'   b  = largura da chaveta / rasgo
'   h  = altura da chaveta
'   t1 = profundidade do rasgo no eixo
'   t2 = profundidade do rasgo no cubo
'==============================================================================
Option Explicit

Public Function BuscarDIN6885(ByVal d As Double, ByRef b As Double, ByRef h As Double, _
                              ByRef t1 As Double, ByRef t2 As Double) As Boolean
    d = Round(d, 3)
    BuscarDIN6885 = True
    Select Case d
        Case Is <= 6:   BuscarDIN6885 = False
        Case Is <= 8:   Definir b, h, t1, t2, 2, 2, 1.2, 1
        Case Is <= 10:  Definir b, h, t1, t2, 3, 3, 1.8, 1.4
        Case Is <= 12:  Definir b, h, t1, t2, 4, 4, 2.5, 1.8
        Case Is <= 17:  Definir b, h, t1, t2, 5, 5, 3, 2.3
        Case Is <= 22:  Definir b, h, t1, t2, 6, 6, 3.5, 2.8
        Case Is <= 30:  Definir b, h, t1, t2, 8, 7, 4, 3.3
        Case Is <= 38:  Definir b, h, t1, t2, 10, 8, 5, 3.3
        Case Is <= 44:  Definir b, h, t1, t2, 12, 8, 5, 3.3
        Case Is <= 50:  Definir b, h, t1, t2, 14, 9, 5.5, 3.8
        Case Is <= 58:  Definir b, h, t1, t2, 16, 10, 6, 4.3
        Case Is <= 65:  Definir b, h, t1, t2, 18, 11, 7, 4.4
        Case Is <= 75:  Definir b, h, t1, t2, 20, 12, 7.5, 4.9
        Case Is <= 85:  Definir b, h, t1, t2, 22, 14, 9, 5.4
        Case Is <= 95:  Definir b, h, t1, t2, 25, 14, 9, 5.4
        Case Is <= 110: Definir b, h, t1, t2, 28, 16, 10, 6.4
        Case Is <= 130: Definir b, h, t1, t2, 32, 18, 11, 7.4
        Case Is <= 150: Definir b, h, t1, t2, 36, 20, 12, 8.4
        Case Is <= 170: Definir b, h, t1, t2, 40, 22, 13, 9.4
        Case Is <= 200: Definir b, h, t1, t2, 45, 25, 15, 10.4
        Case Is <= 230: Definir b, h, t1, t2, 50, 28, 17, 11.4
        Case Else:      BuscarDIN6885 = False
    End Select
End Function

Private Sub Definir(ByRef b As Double, ByRef h As Double, ByRef t1 As Double, ByRef t2 As Double, _
                    ByVal bv As Double, ByVal hv As Double, ByVal t1v As Double, ByVal t2v As Double)
    b = bv: h = hv: t1 = t1v: t2 = t2v
End Sub

' Comprimentos padronizados de chaveta (mm)
Public Function ComprimentosPadrao() As Variant
    ComprimentosPadrao = Array(6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 32, 36, 40, 45, 50, 56, 63, _
                               70, 80, 90, 100, 110, 125, 140, 160, 180, 200, 220, 250, 280, 320, 360, 400)
End Function

' Comprimento padr?o mais pr?ximo de 1,5 x d que caiba no espa?o dispon?vel
' e n?o seja menor que a largura. Retorna 0 se nenhum couber.
Public Function SugerirComprimento(ByVal d As Double, ByVal b As Double, ByVal disponivel As Double) As Double
    Dim v As Variant, i As Long, alvo As Double, melhor As Double
    v = ComprimentosPadrao()
    alvo = 1.5 * d
    For i = LBound(v) To UBound(v)
        If v(i) >= b And v(i) <= disponivel + 0.0001 Then
            If melhor = 0 Or Abs(v(i) - alvo) < Abs(melhor - alvo) Then melhor = v(i)
        End If
    Next
    SugerirComprimento = melhor
End Function
