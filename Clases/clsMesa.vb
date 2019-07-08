Imports Clases.Excepciones

#Region "Clase Mesa"
Namespace MisClases
    Public Class clsMesa 

#Region "Eventos"
        Public Event HanCantadoNumero(ByVal eNumero As Integer)
        Public Event FinDelJuego(ByVal pStrFinal As String)
#End Region

#Region "Propiedades Compartidas"
        Private Shared _linea As Boolean = False
        Public Shared ReadOnly Property Linea() As Boolean
            Get
                Return _linea
            End Get
        End Property
#End Region

#Region "Propiedades"
        Private _numeros As New List(Of Integer)

        Private _jugadores As New List(Of clsJugador)
        Public ReadOnly Property Jugadores() As List(Of clsJugador)
            Get
                Return _jugadores
            End Get
        End Property

        Private _bingo As Boolean = False
        Public ReadOnly Property Bingo() As Boolean
            Get
                Return _bingo
            End Get
        End Property

        Private _strFinal As String = ""
#End Region

#Region "Metodos"
        Public Sub AgregarJugador(ByVal pJugador As clsJugador)

            _jugadores.Add(pJugador)

            AddHandler HanCantadoNumero, AddressOf pJugador.VerificarNumero
            AddHandler pJugador.HanCantadoLinea, AddressOf HanCantadoLinea
            AddHandler pJugador.HanCantadoBingo, AddressOf HanCantadoBingo
        End Sub

        Public Sub CantarNumero()
            Dim rnd As New Random(System.DateTime.Now.Millisecond)
            Dim num As Integer = rnd.Next(1, 90)

            Do While Not (ValidaNumero(num, True))
                num = rnd.Next(1, 90)
            Loop
            _numeros.Add(num)

            RaiseEvent HanCantadoNumero(num)
        End Sub

        Public Sub CantarNumero(ByVal pNumero As Integer)

            If ValidaNumero(pNumero) Then
                _numeros.Add(pNumero)
                RaiseEvent HanCantadoNumero(pNumero)
            End If
        End Sub

        Private Function ValidaNumero(ByVal pNum As Integer, Optional ByVal automatico As Boolean = False) As Boolean

            If pNum < 1 Or pNum > 90 Then
                Dim ex As New Exception

                Dim clsEx As New NumeroInvalido("El numero ingresado esta fuera del rango", ex)
                Throw clsEx
            End If
            If _numeros.Count = 0 Then Return True
            If _numeros.Count = 90 Then
                _strFinal &= "Bingo: Vacante, han sido cantadas todas las bolillas"
                RaiseEvent FinDelJuego(_strFinal)
            End If
            If CBool(_numeros.Find(Function(x) x = pNum)) Then
                If automatico Then Return False
                Dim ex As New Exception

                Dim clsEx As New NumeroRepetido("El numero ingresado ya fue cantado", ex)
                Throw clsEx
            Else
                Return True
            End If
        End Function

        Public Sub HanCantadoBingo(ByVal pJugador As clsJugador)

            _bingo = True
            Dim str As String = vbTab & pJugador.Nombre & " " & (pJugador.NumeroCarton + 1).ToString() & vbCrLf
            If _strFinal = "" Then _strFinal = "Bingo:" & vbCrLf & vbTab & str Else _strFinal &= str
        End Sub

        Public Sub HanCantadoLinea(ByVal pJugador As clsJugador)

            Dim str As String = vbTab & pJugador.Nombre & " " & (pJugador.NumeroCarton + 1).ToString() & vbCrLf
            If _strFinal = "" Then _strFinal = "Linea:" & vbCrLf & vbTab & str Else _strFinal &= str
        End Sub

        Public Sub FinalizarJuego()

            RaiseEvent FinDelJuego(_strFinal)
        End Sub

        Public Sub FinalizoLinea()

            _linea = True
        End Sub

        Public Sub JuegoNuevo()

            For Each _jugador As clsJugador In _jugadores
                RemoveHandler HanCantadoNumero, AddressOf _jugador.VerificarNumero
                RemoveHandler _jugador.HanCantadoLinea, AddressOf HanCantadoLinea
                RemoveHandler _jugador.HanCantadoBingo, AddressOf HanCantadoBingo
            Next
            _jugadores.Clear()
            _numeros.Clear()
            _linea = False
            _bingo = False
            clsCarton.Cantidad = 0
        End Sub
#End Region

    End Class
End Namespace
#End Region

#Region "Excepciones"
Namespace Excepciones

    Public Class NumeroInvalido
        Inherits ApplicationException

        Public Sub New(ByVal pMsg As String, ByVal pInnerEx As Exception)

            MyBase.New(pMsg, pInnerEx)
        End Sub
    End Class

    Public Class NumeroRepetido
        Inherits ApplicationException

        Public Sub New(ByVal pMsg As String, ByVal pInnerEx As Exception)

            MyBase.New(pMsg, pInnerEx)
        End Sub

    End Class
End Namespace
#End Region
