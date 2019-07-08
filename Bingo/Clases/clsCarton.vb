Public Class clsCarton

#Region "Propiedades Compartidas"

    Shared _cantidad As Integer = 0

    Shared ReadOnly Property Cantidad() As Integer
        Get
            Return _cantidad
        End Get

    End Property

#End Region

#Region "Propiedades"

    Public Sub New(ByVal pNombre As String, ByVal pCasillasOcupadas As List(Of Decimal))

        _cantidad += 1
        _nombre = pNombre

        For i As Integer = 1 To 3
            For j As Integer = 1 To 9
                Dim num As Decimal = CDec(i + (j / 10))

                If Not (pCasillasOcupadas.Find(Function(x) x = num)) Then _casillaslibres.Add(num)
            Next
        Next

    End Sub


    Private _nombre As String
    Public ReadOnly Property Nombre() As String
        Get
            Return _nombre
        End Get
    End Property

    Private _casillaslibres As List(Of Decimal)

#End Region

#Region "Metodos"

    Public Function LlenarCarton() As List(Of Integer)

        For Each num As Decimal In _casillaslibres
            Dim aux() As String

            aux = Split(num.ToString())

        Next
    End Function
#End Region
End Class
