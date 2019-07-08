Option Strict Off

Namespace MisClases
    Public Class clsTablero

#Region "Eventos"
        Public Event ActualiazarEstado(ByVal pEvento As Integer, ByVal pDato As String, ByVal pDatoTxt As String)
#End Region

#Region "Propiedades"
        Private _numeros As Integer = 0
        Private _linea As Integer = 0
        Private _bingo As Integer = 0

#End Region

#Region "Metodos"

        Public Sub HanCantadoNumero(ByVal pNumero As Integer)

            _numeros += 1
            RaiseEvent ActualiazarEstado(0, "Bolillas: " & _numeros.ToString() & " / 90 (" & (90 - _numeros).ToString _
                                         & " Restantes)", "Bolilla: " & IIf(pNumero.ToString().Length = 1, "0" _
                                         & pNumero.ToString(), pNumero.ToString()))
        End Sub

        Public Sub HanCantadoLinea(ByVal pJugador As clsJugador)

            _linea += 1
            RaiseEvent ActualiazarEstado(1, "Linea: " & _linea.ToString(), "Linea: " & pJugador.Nombre & " Carton Nro." _
                                         & (pJugador.NumeroCarton + 1).ToString())
        End Sub

        Public Sub HanCantadoBingo(ByVal pJugador As clsJugador)

            _bingo += 1
            RaiseEvent ActualiazarEstado(2, "Bingo: " & _bingo.ToString(), "Bingo: " & pJugador.Nombre & " Carton Nro." _
                                         & (pJugador.NumeroCarton + 1).ToString())
        End Sub

        Public Sub HaFinalizadoElJuego()

            _numeros = 0
            _linea = 0
            _bingo = 0
            RaiseEvent ActualiazarEstado(3, vbCrLf & "Finalizo el Juego", "")
        End Sub
#End Region
    End Class
End Namespace
