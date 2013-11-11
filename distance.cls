Option Explicit
Private Const Pi As Double = 3.14159265358979

Sub Main()
  dblTest = DistanciaDosPuntos("14.637381691871765, -90.51254870069886", "14.637245721896146, -90.51137925756836")
End Sub

Function DistanciaDosPuntos(Optional ByVal strCentroArea As String = "", _
                    Optional ByVal strPunto As String = "") As Double
On Error GoTo ErrorHandler

Dim latlngArea() As String
Dim latlngPunto() As String
Dim lngLatitudArea As Double
Dim lngLongitudArea As Double
Dim lngLatitudPunto As Double
Dim lngLongitudPunto As Double

Dim lngRadioTierra As Double

Dim dLat As Double
Dim dLon As Double
Dim a As Double
Dim c As Double
Dim d As Double

lngRadioTierra = 6378 'kms

latlngArea() = Split(strCentroArea, ",")
latlngPunto() = Split(strPunto, ",")

lngLatitudArea = latlngArea(0)
lngLongitudArea = latlngArea(1)

lngLatitudPunto = latlngPunto(0)
lngLongitudPunto = latlngPunto(1)

dLat = ConvertirARadianes(lngLatitudPunto - lngLatitudArea)
dLon = ConvertirARadianes(lngLongitudPunto - lngLongitudArea)

a = Sin(dLat / 2) * Sin(dLat / 2) + Cos(ConvertirARadianes(lngLatitudArea)) * Cos(ConvertirARadianes(lngLatitudArea)) * Sin(dLon / 2) * Sin(dLon / 2)
c = 2 * Atn2(Sqr(a), Sqr(1 - a))
d = lngRadioTierra * c
d = d * 1000 'metros


DistanciaDosPuntos = d

ErrorHandler:
End Function



'Grados a radianes
Function ConvertirARadianes(Optional ByVal dblValor As Double = -100) As Double
    ConvertirARadianes = dblValor * Pi / 180
End Function

Public Function Atn2(y As Double, x As Double) As Double
  If x > 0 Then
    Atn2 = Atn(y / x)
  ElseIf x < 0 Then
    Atn2 = Sgn(y) * (Pi - Atn(Abs(y / x)))
  ElseIf y = 0 Then
    Atn2 = 0
  Else
    Atn2 = Sgn(y) * Pi / 2
  End If
End Function


