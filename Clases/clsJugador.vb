Imports System.Drawing

Namespace MisClases
    Public Class clsJugador

#Region "Eventos"

        Public Event HanCantadoBingo(ByVal pJugador As clsJugador)
        Public Event HanCantadoLinea(ByVal pJugador As clsJugador)
        Public Event TildarNumero(ByVal pNumero As Integer, pJugador As clsJugador)
#End Region

#Region "Propiedades Compartidas"

        Shared _auxNombre As String
        Shared Property AuxNombre() As String
            Get
                Return _auxNombre
            End Get
            Set(ByVal value As String)
                _auxNombre = value
            End Set
        End Property
#End Region

#Region "Propiedades"

        Private _nombre As String
        Public ReadOnly Property Nombre() As String
            Get
                Return _nombre
            End Get
        End Property

        Private _numeros As List(Of Integer)
        Public ReadOnly Property Numeros() As List(Of Integer)
            Get
                Return _numeros
            End Get
        End Property

        Private _linea0 As Integer
        Private _linea1 As Integer
        Private _linea2 As Integer

        Private _nCarton As Integer
        Public ReadOnly Property NumeroCarton() As Integer
            Get
                Return _nCarton
            End Get
        End Property

        Private _tildeImg As Bitmap
        Public Property TildeImg() As Bitmap
            Get
                Return _tildeImg
            End Get
            Set(ByVal value As Bitmap)
                _tildeImg = value
            End Set
        End Property

        Private _numerosRestantes As Integer = 15

#End Region

#Region "Metodos"

        Public Sub New(ByVal pNombre As String, ByVal pCarton As Integer)

            _nombre = pNombre
            _nCarton = pCarton
        End Sub

        Public Sub New(ByVal pNombre As String, ByVal pCarton As Integer, pNumeros As List(Of Integer))

            _nombre = pNombre
            _numeros = pNumeros
            _nCarton = pCarton
        End Sub

        Public Sub Bingo(ByVal pJugador As clsJugador)

            RaiseEvent HanCantadoBingo(Me)
        End Sub

        Public Sub VerificarNumero(ByVal pNumero As Integer)

            If CBool(_numeros.Find(Function(x) x = pNumero)) Then
                _numerosRestantes -= 1
                RaiseEvent TildarNumero(pNumero, Me)
            End If
            If _numerosRestantes = 0 Then
                RaiseEvent HanCantadoBingo(Me)
            End If
            If (_linea0 = 0 Or _linea1 = 0 Or _linea2 = 0) And Not (clsMesa.Linea) Then
                RaiseEvent HanCantadoLinea(Me)
            End If

        End Sub

        Public Sub CargarNumeros(ByVal lstNumeros As List(Of Integer))

            _numeros = lstNumeros
        End Sub

        Public Sub CargarLineas(ByVal arrayLineas() As Integer)

            _linea0 = arrayLineas(0)
            _linea1 = arrayLineas(1)
            _linea2 = arrayLineas(2)
        End Sub

        Public Sub RestarNumeroLinea(ByVal pLinea As Integer)

            Select Case pLinea
                Case 0
                    _linea0 -= 1
                Case 1
                    _linea1 -= 1
                Case 2
                    _linea2 -= 1
            End Select
        End Sub
#End Region
    End Class
End Namespace
