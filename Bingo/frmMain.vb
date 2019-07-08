Imports Clases.MisClases
Imports Clases.Excepciones
Imports Bingo.My.Resources

Public Class frmMain
    Dim Mesa As clsMesa
    Dim Tablero As clsTablero
    Dim simboloDecimal As String = Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

    Private Sub CargoTablero(ByVal drg As DataGridView)

        With drg
            Dim primerJuego As Boolean = False

            If .Columns.Count = 0 Then primerJuego = True
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False

            For i As Integer = 0 To 9
                If primerJuego Then
                    'Creo las 10 columnas
                    .Columns.Add("col" & i.ToString(), "col" & i.ToString())
                    .Columns(i).Width = 30
                End If
                For j As Integer = 0 To 8
                    If primerJuego Then
                        'Calculo numero
                        Dim num As Integer = (i + 1) + (j * 10)
                        'Agrego fila si no complete las 9
                        If .Rows.Count < 9 Then drg.Rows.Add()
                        'Asigno valor de celda
                        .Rows(j).Cells(i).Value = IIf(num.ToString().Length = 1, "0" & num.ToString(), num.ToString())
                    End If
                    .Rows(j).Cells(i).Style.ForeColor = Color.Gray
                Next
            Next

            If primerJuego Then
                .Height = .Rows(0).Height * 9 + 1
                .Width = .Columns(0).Width * 10 + 1
            End If
            .ClearSelection()
        End With
    End Sub

    Private Sub btnAgregarJugador_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAgregarJugador.Click

        If clsCarton.Cantidad < 15 Then
            If Not (chkAgregarAutomatico.Checked) Then
                Controls.Add(CrearCarton(clsCarton.Cantidad))
                grpControles.SendToBack()
            Else
                Try
                    Dim cantJugadores As Integer = 0
                    Dim okNum As Boolean = True
                    Do
                        Try
                            cantJugadores = CInt(InputBox("Ingesar la cantidad de jugadores a crear"))
                            If cantJugadores + clsCarton.Cantidad > 15 Then
                                MessageBox.Show("Se ha alcanzado el numero maximo de participantes (15)", "Error al crear carton")
                                Exit Sub
                            End If
                            okNum = True
                        Catch ex As InvalidCastException
                            MessageBox.Show("El numero ingresado no es valido", "Numero Invalido")
                            okNum = False
                        Catch ex As Exception
                            MessageBox.Show("El numero ingresado no es valido", "Numero Invalido")
                            okNum = False
                        End Try
                    Loop Until okNum
                    CreaJugadores(clsCarton.Cantidad + cantJugadores - 1, clsCarton.Cantidad)
                    chkAgregarAutomatico.Enabled = False
                    btnAgregarJugador.Enabled = False
                    chkCantarAutomatico.Enabled = True
                    btnJugar.Enabled = True
                Catch ex As Exception
                    MessageBox.Show("El valor ingresado no es valido", "Error datos")
                End Try
            End If
        Else
            MessageBox.Show("Se ha alcanzado el numero maximo de participantes (15)", "Error al crear carton")
        End If
    End Sub

    Private Sub btnCantarNumero_Click(sender As Object, e As EventArgs) Handles btnCantarNumero.Click

        Try
            Dim num As Integer = CInt(InputBox("Ingresar numero"))

            Mesa.CantarNumero(num)
            VerificoMesa()
        Catch ex As InvalidCastException
            MessageBox.Show("El numero ingresado no es valido", "Error de datos")
        Catch ex As NumeroInvalido
            MessageBox.Show(ex.Message, "Error datos")
        Catch ex As NumeroRepetido
            MessageBox.Show(ex.Message, "Error datos")
        End Try
    End Sub

    Public Function CrearCarton(ByVal pNumero As Integer, Optional ByVal pJugador As String = "",
                                Optional ByVal lstLbl As List(Of Label) = Nothing) As GroupBox

        Dim grp As New GroupBox()
        Dim nombreJugador As String
        Dim strName As String = "grpCarton" & pNumero.ToString()
        If String.IsNullOrEmpty(pJugador) Then
            Do
                nombreJugador = InputBox("Ingresar nombre del jugador", "", " ")
                If nombreJugador = "" Then Exit Function
            Loop Until Not (String.IsNullOrEmpty(nombreJugador.Trim()))
            clsJugador.AuxNombre = nombreJugador
        Else : nombreJugador = pJugador
        End If

        If String.IsNullOrEmpty(pJugador) Then strName &= "A"
        grp.Name = strName
        grp.Text = nombreJugador & " - Carton Nro." & (pNumero + 1).ToString()
        grp.Tag = pNumero.ToString()
        If String.IsNullOrEmpty(pJugador) Then
            grp.Width = 314
            grp.Height = 272
            grp.Left = grpControles.Left
            grp.Top = grpControles.Top
        Else
            grp.Width = 308
            grp.Height = 123
            If pNumero = 0 Or pNumero = 3 Or pNumero = 6 Or pNumero = 9 Or pNumero = 12 Then grp.Left = 10
            If pNumero = 1 Or pNumero = 4 Or pNumero = 7 Or pNumero = 10 Or pNumero = 13 Then grp.Left = grp.Width + 20
            If pNumero = 2 Or pNumero = 5 Or pNumero = 8 Or pNumero = 11 Or pNumero = 14 Then grp.Left = grp.Width * 2 + 30
            If pNumero = 0 Or pNumero = 1 Or pNumero = 2 Then grp.Top = 12
            If pNumero = 3 Or pNumero = 4 Or pNumero = 5 Then grp.Top = grp.Height + 24
            If pNumero = 6 Or pNumero = 7 Or pNumero = 8 Then grp.Top = grp.Height * 2 + 36
            If pNumero = 9 Or pNumero = 10 Or pNumero = 11 Then grp.Top = grp.Height * 3 + 48
            If pNumero = 12 Or pNumero = 13 Or pNumero = 14 Then grp.Top = grp.Height * 4 + 60
        End If
        'Armo casilleros
        For i As Integer = 0 To 2
            For j As Integer = 0 To 8
                Dim nlabel As New Label()

                nlabel.Name = "lbl" & pNumero.ToString() & i.ToString() & j.ToString()
                nlabel.AutoSize = False
                nlabel.Width = 32
                nlabel.Height = 32
                nlabel.Top = (i * 32) + 16
                nlabel.Left = (j * 32) + 13
                nlabel.Font = New Font("Microsoft Sans Serif", 11, FontStyle.Bold)
                nlabel.BorderStyle = BorderStyle.FixedSingle
                nlabel.TextAlign = ContentAlignment.MiddleCenter
                nlabel.ImageAlign = ContentAlignment.MiddleCenter
                nlabel.BackColor = Color.White
                If String.IsNullOrEmpty(pJugador) Then
                    AddHandler nlabel.Click, AddressOf MarcarCasilla
                Else
                    Dim newLbl As Label = lstLbl.Find(Function(x) x.Name.Contains(nlabel.Name))
                    If Not (IsNothing(newLbl)) Then
                        nlabel.Text = newLbl.Text
                        nlabel.BackColor = newLbl.BackColor
                    End If
                End If
                grp.Controls.Add(nlabel)
            Next
        Next
        If Not (String.IsNullOrEmpty(pJugador)) Then
            grp.BringToFront()
            Return grp
        End If
        'Armo Controles configuracion carton
        Dim lbl As New Label()
        lbl.Name = "lblSelecTilde"
        lbl.AutoSize = True
        lbl.Left = 6
        lbl.Top = 158
        lbl.Font = New Font("Microsoft Sans Serif", 10, FontStyle.Bold)
        lbl.TextAlign = ContentAlignment.MiddleCenter
        lbl.Text = "Seleccionar Tilde"
        grp.Controls.Add(lbl)

        Dim lblContador As New Label()
        lblContador.Name = "lblContador"
        lblContador.Text = "01/12"
        lblContador.AutoSize = True
        lblContador.Left = 21
        lblContador.Top = 118
        grp.Controls.Add(lblContador)

        For i As Integer = 0 To 4
            Dim nlbl As New Label()
            nlbl.Name = "lblImg" & i.ToString()
            nlbl.Text = ""
            nlbl.TextAlign = ContentAlignment.MiddleCenter
            nlbl.ImageAlign = ContentAlignment.MiddleCenter
            nlbl.Width = 32
            nlbl.Height = 32
            nlbl.Top = 183
            nlbl.BorderStyle = BorderStyle.None
            Select Case i
                Case 0
                    nlbl.Image = Tilde
                    nlbl.Left = 29
                Case 1
                    nlbl.Image = Tilde1
                    nlbl.Left = 86
                Case 2
                    nlbl.Image = Tilde2
                    nlbl.Left = 143
                Case 3
                    nlbl.Image = Tilde3
                    nlbl.Left = 200
                Case 4
                    nlbl.Image = Tilde4
                    nlbl.Left = 257
            End Select
            AddHandler nlbl.Click, AddressOf SeleccionarImg
            grp.Controls.Add(nlbl)
        Next

        Dim chk As New CheckBox()
        chk.Name = "chkArmaNumeros"
        chk.Tag = nombreJugador
        chk.Text = "Llenar Casilleros"
        chk.Width = 102
        chk.Height = 17
        chk.Left = 102
        chk.Top = 122
        grp.Controls.Add(chk)

        Dim btn As New Button()
        btn.Name = "btnColores"
        btn.Text = "Colores"
        btn.Width = 54
        btn.Height = 23
        btn.Left = 247
        btn.Top = 118
        AddHandler btn.Click, AddressOf SeleccionarColor
        grp.Controls.Add(btn)

        Dim btnAceptar As New Button()
        btnAceptar.Name = "btnAceptar"
        btnAceptar.Text = "Aceptar"
        btnAceptar.Tag = pNumero.ToString()
        btnAceptar.Width = 75
        btnAceptar.Height = 23
        btnAceptar.Left = 6
        btnAceptar.Top = 238
        AddHandler btnAceptar.Click, AddressOf AceptarCarton
        grp.Controls.Add(btnAceptar)

        Dim btnCancelar As New Button()
        btnCancelar.Name = "btnCancelar"
        btnCancelar.Text = "Cancelar"
        btnCancelar.Tag = pNumero.ToString()
        btnCancelar.Width = btnAceptar.Width
        btnCancelar.Top = btnAceptar.Height
        btnCancelar.Left = btnAceptar.Left
        btnCancelar.Left = 233
        btnCancelar.Top = 238
        AddHandler btnCancelar.Click, AddressOf CancelarCarton
        grp.Controls.Add(btnCancelar)
        grp.BringToFront()

        Return grp
    End Function

    Private Sub EliminarGrp(ByVal pNombre As String)

        Dim ctrls() As Control = Controls.Find("grpCarton" & pNombre, True)

        For Each ctrl As Control In ctrls
            Controls.Remove(ctrl)
        Next
    End Sub

#Region "Handles"
    Private Sub MarcarTablero(ByVal pNum As Integer)

        If pNum = -1 Then
            MessageBox.Show("El numero ya fue cantado", "Error Numero Ingresado")
            Exit Sub
        End If
        Dim pos() As String = Split(CStr(pNum / 10), simboloDecimal)
        Dim x As Integer, y As Integer

        If pos.Count = 1 Then
            y = 9
            x = CInt(pos(0)) - 1
        Else
            x = CInt(pos(0))
            y = CInt(pos(1) - 1)
        End If

        drgTablero.Rows(x).Cells(y).Style.ForeColor = Color.Red
    End Sub

    Private Sub MarcarCarton(ByVal pNum As Integer, ByVal pJugador As clsJugador)
        Dim pos() As String = Split(CStr(pNum / 10), simboloDecimal)
        Dim x As Integer

        x = CInt(pos(0))
        If x = 9 Then x = 8

        For i As Integer = 0 To 2
            Dim keyCtrl As String = "lbl" & pJugador.NumeroCarton & i.ToString() & x.ToString()
            Dim ctrls() As Control = Controls.Find(keyCtrl, True)
            For Each ctrl As Label In ctrls
                If ctrl.BackColor = Color.White Then
                    If CInt(ctrl.Text) = pNum Then
                        ctrl.Image = pJugador.TildeImg
                        pJugador.RestarNumeroLinea(i)
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub MarcarCasilla(sender As Object, e As EventArgs)
        Dim lbl As Label = sender
        Dim col As Integer = Mid(lbl.Name, lbl.Name.Length)
        Dim casillasOcupadas As Integer = 0

        If lbl.BackColor = Color.White Then
            Dim str As String = clsCarton.MarcarCasilla(col, True, casillasOcupadas)

            If String.IsNullOrEmpty(str) Then lbl.BackColor = dlgColor.Color Else MessageBox.Show(str, "Error creacion carton")
        Else
            Dim str As String = clsCarton.MarcarCasilla(col, False, casillasOcupadas)

            If String.IsNullOrEmpty(str) Then lbl.BackColor = Color.White Else MessageBox.Show(str, "Error creacion carton")
        End If
        Dim ctrls() As Control = Controls.Find("lblContador", True)
        For Each ctrl As Control In ctrls
            ctrl.Text = IIf(casillasOcupadas.ToString().Length = 1, "0" & casillasOcupadas, casillasOcupadas) & "/ 12"
        Next
    End Sub

    Private Sub SeleccionarImg(sender As Object, e As EventArgs)
        Dim lbl As Label = sender
        lbl.BorderStyle = BorderStyle.FixedSingle

        For i As Integer = 0 To 4
            Dim ctrls() As Control = Controls.Find("lblImg" & i.ToString(), True)
            For Each ctrl As Label In ctrls
                If ctrl.Name <> lbl.Name Then ctrl.BorderStyle = BorderStyle.None
            Next
        Next
    End Sub

    Private Sub SeleccionarColor(sender As Object, e As EventArgs)

        dlgColor.ShowDialog()
    End Sub

    Private Sub CancelarCarton(sender As Object, e As EventArgs)
        Dim btn As Button = sender

        EliminarGrp(btn.Tag & "A")
    End Sub

    Private Sub AceptarCarton(sender As Object, e As EventArgs)
        Dim btn As Button = sender
        Dim lstLabels As New List(Of Label)
        Dim lstNumeros As New List(Of Integer)
        Dim tildeImg As Bitmap = Nothing
        Dim grpCarton As GroupBox = CType(Controls("grpCarton" & btn.Tag.ToString() & "A"), GroupBox)

        If clsCarton.CasillasOcupadas < 12 Then
            MessageBox.Show("Faltan casillas por ocupar" & vbCrLf _
                & "15 Casillas Disponibles / 12 Casillas Pintadas", "Error creacion carton")
            Exit Sub
        End If
        'Valido que se haya seleccionado alguna imagen para tildar los numeros
        For i As Integer = 0 To 4
            With CType(grpCarton.Controls("lblImg" & i.ToString()), Label)
                If .BorderStyle = BorderStyle.FixedSingle Then
                    tildeImg = .Image
                    Exit For
                End If
            End With
        Next
        If IsNothing(tildeImg) Then
            MessageBox.Show("Debe seleccionar el tilde para el casillero", "Error creacion carton")
            Exit Sub
        End If
        'Quito Handler a los labels
        'Armo lista con los casilleros ocupados
        Dim NombreJugador As String = clsJugador.AuxNombre
        Dim pCarton As New clsCarton(NombreJugador)
        'Creo la instancia del nuevo jugador
        Dim pJugador As clsJugador
        'Verifico si genero numeros automaticamente
        Dim generoNumeros As Boolean = False
        If CType(grpCarton.Controls("chkArmaNumeros"), CheckBox).Checked Then generoNumeros = True
        Dim arrayLineas() As Integer = {0, 0, 0}
        For i As Integer = 0 To 2
            For j As Integer = 0 To 8
                Dim ctrls() As Control = Controls.Find("lbl" & btn.Tag.ToString() & i.ToString() & j.ToString(), True)
                For Each ctrl As Label In ctrls
                    RemoveHandler ctrl.Click, AddressOf MarcarCasilla
                    lstLabels.Add(ctrl)
                    Dim numero As Integer
                    If ctrl.BackColor = Color.White Then
                        Dim ok As Boolean = False
                        If generoNumeros Then
                            Do
                                numero = pCarton.ObtenerNumero(j)
                                If lstNumeros.Find(Function(x) x = numero) Then ok = False Else ok = True
                            Loop Until ok
                        Else
                            Do
                                Dim okNum As Boolean = True
                                Do
                                    Try
                                        numero = CInt(InputBox("Ingresar numero para la casilla: " _
                                            & (i + 1).ToString & "," & (j + 1).ToString() & vbCrLf _
                                            & ObtenerRangoCasilla(j)))
                                        okNum = True
                                    Catch ex As InvalidCastException
                                        MessageBox.Show("El numero ingresado no es valido", "Numero Invalido")
                                        okNum = False
                                    Catch ex As Exception
                                        MessageBox.Show("El numero ingresado no es valido", "Numero Invalido")
                                        okNum = False
                                    End Try
                                Loop Until okNum
                                Dim strAux As String = pCarton.ObtenerNumero(j, numero)
                                If Not (String.IsNullOrEmpty(strAux)) Then MessageBox.Show(strAux, "Numero Invalido") Else ok = True
                                If lstNumeros.Find(Function(x) x = numero) Then
                                    MessageBox.Show("El numero ingresado ya fue cargado", "Numero Repetido")
                                    ok = False
                                End If
                            Loop Until ok
                        End If
                        ctrl.Text = IIf(numero.ToString().Length = 1, "0" & numero.ToString(), numero.ToString())
                        lstNumeros.Add(numero)
                        arrayLineas(i) += 1
                    End If
                Next
            Next
        Next
        'Agrego jugador a la mesa
        If generoNumeros Then
            pJugador = New clsJugador(NombreJugador, btn.Tag, lstNumeros)
        Else
            pJugador = New clsJugador(NombreJugador, btn.Tag)
            pJugador.CargarNumeros(lstNumeros)
        End If
        pJugador.TildeImg = tildeImg
        pJugador.CargarLineas(arrayLineas)
        AddHandler pJugador.HanCantadoLinea, AddressOf Linea
        AddHandler pJugador.HanCantadoBingo, AddressOf Bingo
        AddHandler pJugador.HanCantadoLinea, AddressOf Tablero.HanCantadoLinea
        AddHandler pJugador.HanCantadoBingo, AddressOf Tablero.HanCantadoBingo
        AddHandler pJugador.TildarNumero, AddressOf MarcarCarton
        Mesa.AgregarJugador(pJugador)
        'Genero el carton en la mesa
        Controls.Add(CrearCarton(btn.Tag, NombreJugador, lstLabels))
        'Elimino el control recien creado para configurar el carton y limpo el listado de labels
        lstLabels.Clear()
        EliminarGrp(btn.Tag & "A")
        btnJugar.Enabled = True
        chkCantarAutomatico.Enabled = True

    End Sub

    Private Sub Bingo(ByVal pJugador As clsJugador)

        timerCantar.Enabled = False
        MessageBox.Show("Bingo!! " & vbCrLf & "Ganador Jugador: " & pJugador.Nombre & vbCrLf & "Carton: " _
                           & (pJugador.NumeroCarton + 1).ToString(), "Jugador Canto Bingo")
    End Sub

    Private Sub Linea(ByVal pJugador As clsJugador)

        If Not (chkCantarAutomatico.Checked) Then MessageBox.Show("Linea!! " & vbCrLf & "Jugador: " _
            & pJugador.Nombre & vbCrLf & "Carton: " & (pJugador.NumeroCarton + 1).ToString(), "Jugador Canto Linea")
        Mesa.FinalizoLinea()
    End Sub

    Private Sub ActualizaEstado(ByVal pEvento As Integer, ByVal pDato As String, ByVal pDatoTxt As String)

        If pEvento <> 3 Then
            'Dim ctrls() As Control = grpEstadoDelJuego.Controls.Find("lblEvento" & pEvento.ToString(), True)
            Dim lbl As Label = CType(grpEstadoDelJuego.Controls("lblEvento" & pEvento.ToString()), Label)
            'For Each ctrl As Label In ctrls
            lbl.Text = pDato
            txtEventosJuego.Text = pDatoTxt & vbCrLf & txtEventosJuego.Text
                If pEvento = 0 Then
                    Dim num As String = Trim(Mid(pDatoTxt, InStr(pDatoTxt, " ")))
                    lblNumeroCantado.Text = num
                End If
            'Next
        Else : txtEventosJuego.Text = pDato & vbCrLf & txtEventosJuego.Text
        End If
    End Sub
#End Region

    Private Sub btnJugar_Click(sender As Object, e As EventArgs) Handles btnJugar.Click

        btnAgregarJugador.Enabled = False
        chkCantarAutomatico.Enabled = False
        chkAgregarAutomatico.Enabled = False
        If chkCantarAutomatico.Checked Then
            btnCantarNumero.Enabled = False
            timerCantar.Interval = 1250
            timerCantar.Enabled = True
        Else
            btnCantarNumero.Enabled = True
        End If
        txtEventosJuego.Text = vbCrLf & "Comienzo del Juego"
        btnJugar.Enabled = False
    End Sub

    Private Sub btnJuegoNuevo_Click(sender As Object, e As EventArgs) Handles btnJuegoNuevo.Click

        If IsNothing(Tablero) Then Tablero = New clsTablero()
        If IsNothing(Mesa) Then
            Mesa = New clsMesa()
            AddHandler Mesa.HanCantadoNumero, AddressOf MarcarTablero
            AddHandler Mesa.HanCantadoNumero, AddressOf Tablero.HanCantadoNumero
            AddHandler Mesa.FinDelJuego, AddressOf Tablero.HaFinalizadoElJuego
            AddHandler Tablero.ActualiazarEstado, AddressOf ActualizaEstado
        Else
            Mesa.JuegoNuevo()
        End If

        LimpioTablero()
        CargoTablero(drgTablero)
    End Sub

    Private Sub timerCantar_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles timerCantar.Tick

        Mesa.CantarNumero()
        VerificoMesa(True)
    End Sub

    Private Sub VerificoMesa(Optional ByVal automatico As Boolean = False)

        If Mesa.Bingo Then
            If automatico Then timerCantar.Enabled = False
            Mesa.FinalizarJuego()
            btnCantarNumero.Enabled = False
            btnJuegoNuevo.Enabled = True
        End If
    End Sub

    Private Sub CreaJugadores(ByVal pCantidad As Integer, Optional pInicio As Integer = 0)
        Dim rnd As New Random(Date.Now.Millisecond)
        Dim lstNumeros As List(Of Integer)
        Dim lstPos As New List(Of String)

        For k As Integer = pInicio To pCantidad
            Dim Nombre As String = ""
            Do
                Nombre = InputBox("Ingresar nombre jugador nro. " & (k + 1).ToString())
            Loop Until Not (String.IsNullOrEmpty(Nombre))

            Dim grp As New GroupBox
            grp.Name = "grpCarton" & k.ToString()
            grp.Text = Nombre & " - Carton Nro." & (k + 1).ToString()
            grp.Tag = k.ToString()
            grp.Width = 308
            grp.Height = 123
            If k = 0 Or k = 3 Or k = 6 Or k = 9 Or k = 12 Then grp.Left = 10
            If k = 1 Or k = 4 Or k = 7 Or k = 10 Or k = 13 Then grp.Left = grp.Width + 20
            If k = 2 Or k = 5 Or k = 8 Or k = 11 Or k = 14 Then grp.Left = grp.Width * 2 + 30
            If k = 0 Or k = 1 Or k = 2 Then grp.Top = 12
            If k = 3 Or k = 4 Or k = 5 Then grp.Top = grp.Height + 24
            If k = 6 Or k = 7 Or k = 8 Then grp.Top = grp.Height * 2 + 36
            If k = 9 Or k = 10 Or k = 11 Then grp.Top = grp.Height * 3 + 48
            If k = 12 Or k = 13 Or k = 14 Then grp.Top = grp.Height * 4 + 60
            'Armo casilleros
            lstNumeros = New List(Of Integer)
            Dim arrayLineas() As Integer = {0, 0, 0}
            Dim pCarton As New clsCarton(Nombre)
            For i As Integer = 0 To 2
                For j As Integer = 0 To 8
                    Dim nlabel As New Label()

                    nlabel.Name = "lbl" & k.ToString() & i.ToString() & j.ToString()
                    nlabel.AutoSize = False
                    nlabel.Width = 32
                    nlabel.Height = 32
                    nlabel.Top = (i * 32) + 16
                    nlabel.Left = (j * 32) + 9
                    nlabel.Font = New Font("Microsoft Sans Serif", 11, FontStyle.Bold)
                    nlabel.BorderStyle = BorderStyle.FixedSingle
                    nlabel.TextAlign = ContentAlignment.MiddleCenter
                    nlabel.ImageAlign = ContentAlignment.MiddleCenter
                    nlabel.BackColor = Color.White
                    Dim ok As Boolean = False
                    Dim numero As Integer
                    Do
                        numero = pCarton.ObtenerNumero(j)
                        If lstNumeros.Find(Function(x) x = numero) Then ok = False Else ok = True
                    Loop Until ok
                    nlabel.Text = IIf(numero.ToString().Length = 1, "0" & numero.ToString(), numero.ToString())
                    lstNumeros.Add(numero)
                    arrayLineas(i) += 1
                    grp.Controls.Add(nlabel)
                Next
            Next
            'Armo color aleatorio para las casillas
            Dim newColor As Color = Color.FromArgb(rnd.Next(0, 255), rnd.Next(0, 255), rnd.Next(0, 255))
            'Selecciono tilde aleatorio para las coincidencias
            Dim nTilde As Integer = rnd.Next(0, 4)
            Dim imgTilde As Bitmap = Nothing
            Select Case nTilde
                Case 0
                    imgTilde = Tilde
                Case 1
                    imgTilde = Tilde1
                Case 2
                    imgTilde = Tilde2
                Case 3
                    imgTilde = Tilde3
                Case 4
                    imgTilde = Tilde4
            End Select
            'Armo casillas aleatorias
            lstPos.Clear()
            For i As Integer = 0 To 11
                Dim x As Integer
                Dim y As Integer
                Dim ok As Boolean = False
                Do
                    y = rnd.Next(0, 2)
                    x = rnd.Next(0, 8)
                    Dim strPos As String = lstPos.Find(Function(p) p.Contains(y.ToString() & "," & x.ToString()))
                    If strPos = Nothing Then ok = True
                Loop Until ok

                Dim ctrls() As Control = grp.Controls.Find("lbl" & k.ToString() & y.ToString() & x.ToString(), True)
                For Each ctrl As Label In ctrls
                    lstNumeros.Remove(CInt(ctrl.Text))
                    ctrl.Text = ""
                    ctrl.BackColor = newColor
                    arrayLineas(y) -= 1
                    lstPos.Add(y.ToString() & "," & x.ToString())
                Next
            Next
            Dim pJugador As New clsJugador(Nombre, k)
            pJugador.CargarNumeros(lstNumeros)
            pJugador.TildeImg = imgTilde
            pJugador.CargarLineas(arrayLineas)
            AddHandler pJugador.HanCantadoLinea, AddressOf Linea
            AddHandler pJugador.HanCantadoBingo, AddressOf Bingo
            AddHandler pJugador.HanCantadoLinea, AddressOf Tablero.HanCantadoLinea
            AddHandler pJugador.HanCantadoBingo, AddressOf Tablero.HanCantadoBingo
            AddHandler pJugador.TildarNumero, AddressOf MarcarCarton
            Mesa.AgregarJugador(pJugador)
            'Genero el carton en la mesa
            Controls.Add(grp)
        Next
    End Sub

    Private Sub LimpioTablero()

        For Each grpCarton As GroupBox In (From g In Controls.OfType(Of GroupBox)
                                           Where g.Name.Contains("grpCarton")
                                           Select g).ToList()
            Controls.Remove(grpCarton)
            grpCarton.Dispose()
        Next

        chkAgregarAutomatico.Checked = False
        chkCantarAutomatico.Checked = False
        btnAgregarJugador.Enabled = True
        chkAgregarAutomatico.Enabled = True
        btnJuegoNuevo.Enabled = False
        lblEvento0.Text = "Bolillas: 0 / 90 (Restantes 90)"
        lblEvento1.Text = "Linea: 0"
        lblEvento2.Text = "Bingo: 0"
        txtEventosJuego.Clear()
    End Sub

    Private Function ObtenerRangoCasilla(ByVal pCasilla As Integer) As String
        Dim numero As String = ""

        Select Case pCasilla
            Case 0
                numero = "Rango 1 - 9"
            Case 1
                numero = "Rango 10 - 19"
            Case 2
                numero = "Rango 20 - 29"
            Case 3
                numero = "Rango 30 - 39"
            Case 4
                numero = "Rango 40 - 49"
            Case 5
                numero = "Rango 50 - 59"
            Case 6
                numero = "Rango 60 - 69"
            Case 7
                numero = "Rango 70 - 79"
            Case 8
                numero = "Rango 80 - 90"
        End Select

        Return numero
    End Function
End Class
