Namespace MisClases
    Public Class clsCarton

#Region "Propiedades Compartidas"

        Shared _cantidad As Integer = 0
        Shared Property Cantidad() As Integer
            Get
                Return _cantidad
            End Get
            Friend Set(ByVal value As Integer)
                _cantidad = value
            End Set
        End Property

        Private Shared _casillasocupadas As Integer
        Shared Property CasillasOcupadas() As Integer
            Get
                Return _casillasocupadas
            End Get
            Private Set(ByVal value As Integer)
                _casillasocupadas = value
            End Set
        End Property

        Shared _col0 As Integer
        Shared Property Col0() As Integer
            Get
                Return _col0
            End Get
            Private Set(ByVal value As Integer)
                _col0 = value
            End Set
        End Property

        Shared _col1 As Integer
        Shared Property Col1() As Integer
            Get
                Return _col1
            End Get
            Private Set(ByVal value As Integer)
                _col1 = value
            End Set
        End Property

        Shared _col2 As Integer
        Shared Property Col2() As Integer
            Get
                Return _col2
            End Get
            Private Set(ByVal value As Integer)
                _col2 = value
            End Set
        End Property

        Shared _col3 As Integer
        Shared Property Col3() As Integer
            Get
                Return _col3
            End Get
            Private Set(ByVal value As Integer)
                _col3 = value
            End Set
        End Property

        Shared _col4 As Integer
        Shared Property Col4() As Integer
            Get
                Return _col4
            End Get
            Private Set(ByVal value As Integer)
                _col4 = value
            End Set
        End Property

        Shared _col5 As Integer
        Shared Property Col5() As Integer
            Get
                Return _col5
            End Get
            Private Set(ByVal value As Integer)
                _col5 = value
            End Set
        End Property

        Shared _col6 As Integer
        Shared Property Col6() As Integer
            Get
                Return _col6
            End Get
            Private Set(ByVal value As Integer)
                _col6 = value
            End Set
        End Property

        Shared _col7 As Integer
        Shared Property Col7() As Integer
            Get
                Return _col7
            End Get
            Private Set(ByVal value As Integer)
                _col7 = value
            End Set
        End Property

        Shared _col8 As Integer
        Shared Property Col8() As Integer
            Get
                Return _col8
            End Get
            Private Set(ByVal value As Integer)
                _col8 = value
            End Set
        End Property

#End Region

#Region "Propiedades"

        Private _linea0 As Integer = 0
        Private _linea1 As Integer = 0
        Private _linea2 As Integer = 0

#End Region

#Region "Metodos Compartidos"

        ''' <summary>
        ''' Funcion para marcar casilla del carton 
        ''' </summary>
        ''' <param name="pCol">Indica numero de columna sobre el que se hizo click</param>
        ''' <param name="pAccion">Indica si se tildo (1) o destildo(0)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Shared Function MarcarCasilla(ByVal pCol As Integer, pAccion As Boolean, ByRef pCasillasOcupadas As Integer) As String
            Dim errorColumna As String = ""
            Dim errorCarton As String = ""

            Select Case pCol
                Case 0
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col0 < 2 Then
                                _col0 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col0 -= 1
                    End If
                Case 1
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col1 < 2 Then
                                _col1 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col1 -= 1
                    End If
                Case 2
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col2 < 2 Then
                                _col2 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col2 -= 1
                    End If
                Case 3
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col3 < 2 Then
                                _col3 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col3 -= 1
                    End If
                Case 4
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col4 < 2 Then
                                _col4 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col4 -= 1
                    End If
                Case 5
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col5 < 2 Then
                                _col5 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col5 -= 1
                    End If
                Case 6
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col6 < 2 Then
                                _col6 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col6 -= 1
                    End If
                Case 7
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col7 < 2 Then
                                _col7 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col7 -= 1
                    End If
                Case 8
                    If pAccion Then
                        If CasillasOcupadas < 12 Then
                            If _col8 < 2 Then
                                _col8 += 1
                            Else
                                errorColumna = "No se pueden ocupar todas las casillas de la columna"
                            End If
                        Else
                            errorColumna = "No pueden ocuparse mas de 12 casillas del carton"
                        End If
                    Else
                        _col8 -= 1
                    End If
            End Select
            If pAccion Then
                If String.IsNullOrEmpty(errorColumna) And String.IsNullOrEmpty(errorCarton) Then _casillasocupadas += 1
            Else
                _casillasocupadas -= 1
            End If
            pCasillasOcupadas = _casillasocupadas

            Return errorColumna
        End Function
#End Region

#Region "Metodos"

        Public Sub New(ByVal pNombre As String)

            _cantidad += 1
            _col0 = 0
            _col1 = 0
            _col2 = 0
            _col3 = 0
            _col4 = 0
            _col5 = 0
            _col6 = 0
            _col7 = 0
            _col8 = 0
            _casillasocupadas = 0
        End Sub

        Public Function ObtenerNumero(ByVal pCasilla As Integer) As Integer
            'Pauso 30 ms para obtener un mejor random
            Threading.Thread.Sleep(30)
            Dim rnd As New Random(Date.Now.Millisecond)
            Dim num As Integer

            Select Case pCasilla
                Case 0
                    num = rnd.Next(1, 9)
                Case 1
                    num = rnd.Next(10, 19)
                Case 2
                    num = rnd.Next(20, 29)
                Case 3
                    num = rnd.Next(30, 39)
                Case 4
                    num = rnd.Next(40, 49)
                Case 5
                    num = rnd.Next(50, 59)
                Case 6
                    num = rnd.Next(60, 69)
                Case 7
                    num = rnd.Next(70, 79)
                Case 8
                    num = rnd.Next(80, 90)
            End Select

            Return num
        End Function

        Public Function ObtenerNumero(ByVal pCasilla As Integer, ByVal pNumero As Integer) As String
            Dim numero As String = ""

            Select Case pCasilla
                Case 0
                    If pNumero < 1 Or pNumero > 9 Then numero = "Numero invalido. Rango 1 - 9"
                Case 1
                    If pNumero < 10 Or pNumero > 19 Then numero = "Numero invalido. Rango 10 - 19"
                Case 2
                    If pNumero < 20 Or pNumero > 29 Then numero = "Numero invalido. Rango 20 - 29"
                Case 3
                    If pNumero < 30 Or pNumero > 39 Then numero = "Numero invalido. Rango 30 - 39"
                Case 4
                    If pNumero < 40 Or pNumero > 49 Then numero = "Numero invalido. Rango 40 - 49"
                Case 5
                    If pNumero < 50 Or pNumero > 59 Then numero = "Numero invalido. Rango 50 - 59"
                Case 6
                    If pNumero < 60 Or pNumero > 69 Then numero = "Numero invalido. Rango 60 - 69"
                Case 7
                    If pNumero < 70 Or pNumero > 79 Then numero = "Numero invalido. Rango 70 - 79"
                Case 8
                    If pNumero < 80 Or pNumero > 90 Then numero = "Numero invalido. Rango 80 - 90"
            End Select

            Return numero
        End Function
#End Region
    End Class
End Namespace
