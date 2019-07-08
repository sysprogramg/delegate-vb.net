Public Class Form1


    Public Event CargoNumeroLbl(ByVal pNombre As String, ByVal pNumLbl As String)

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        CargoTablero(Me.drgTablero)
    End Sub

    Private Sub CargoTablero(ByVal drg As DataGridView)

        drg.ColumnHeadersVisible = False
        drg.RowHeadersVisible = False
        drg.AllowUserToResizeColumns = False
        drg.AllowUserToResizeRows = False

        For i As Integer = 0 To 9
            'Creo las 10 columnas
            drg.Columns.Add("col" & i.ToString(), "col" & i.ToString())
            drg.Columns(i).Width = 30
            For j As Integer = 0 To 8
                'Calculo numero
                Dim num As Integer = (i + 1) + (j * 10)
                'Agrego fila si no complete las 9
                If drg.Rows.Count < 9 Then drg.Rows.Add()
                'Asigno valor de celda
                drg.Rows(j).Cells(i).Value = IIf(num.ToString().Length = 1, "0" & num.ToString(), num.ToString())
                drg.Rows(j).Cells(i).Style.ForeColor = Color.Gray
            Next
        Next

        drg.Height = drg.Rows(0).Height * 9 + 1
        drg.Width = drg.Columns(0).Width * 10 + 1
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles btnAgregarJugador.Click
        Dim nJugador As Integer = 1, num As Integer
        Dim grp As New GroupBox()
        Dim rnd As New Random(System.DateTime.Now.Millisecond)
        Dim lstNumeros As New List(Of Integer)


        grp.Name = "grpJugador" & nJugador.ToString()
        grp.Text = "Jugador " & nJugador.ToString()
        grp.Visible = True

        For i As Integer = 0 To 2
            For j As Integer = 0 To 8
                Dim nlabel As New Label()

                nlabel.Name = "lblJugador" & nJugador.ToString() & i.ToString()
                nlabel.AutoSize = False
                nlabel.Width = 32
                nlabel.Height = 32
                nlabel.Top = (i * 32) + 16
                nlabel.Left = (j * 32) + 9
                nlabel.Font = New Font("Microsoft Sans Serif", 11, FontStyle.Bold)
                nlabel.BorderStyle = BorderStyle.FixedSingle
                nlabel.TextAlign = ContentAlignment.MiddleCenter
                nlabel.BackColor = Color.White

                Select Case j
                    Case 0
                        If i = 0 Or i = 1 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(1, 9)
                            lstNumeros.Add(num)
                            nlabel.Text = "0" & num.ToString()
                        End If
                    Case 1
                        If i = 2 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(10, 19)
                            lstNumeros.Add(num)
                            nlabel.Text = num.ToString()
                        End If
                    Case 2
                        If i = 0 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(20, 29)
                            lstNumeros.Add(num)
                            nlabel.Text = num.ToString()
                        End If
                    Case 3
                        If i = 1 Or i = 2 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(30, 39)
                            lstNumeros.Add(num)
                            nlabel.Text = num.ToString()
                        End If
                    Case 4
                        If i = 2 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(40, 49)
                            lstNumeros.Add(num)
                            nlabel.Text = num.ToString()
                        End If
                    Case 5
                        If i = 0 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(50, 59)
                            lstNumeros.Add(num)
                            nlabel.Text = num.ToString()
                        End If
                    Case 6
                        If i = 1 Or i = 2 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(60, 69)
                            lstNumeros.Add(num)
                            nlabel.Text = num.ToString()
                        End If
                    Case 7
                        If i = 0 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(70, 79)
                            lstNumeros.Add(num)
                            nlabel.Text = num.ToString()
                        End If
                    Case 8
                        If i = 1 Then
                            nlabel.BackColor = Color.Black
                        Else
                            num = rnd.Next(80, 90)
                            lstNumeros.Add(num)
                            nlabel.Text = num.ToString()
                        End If
                End Select
                grp.Controls.Add(nlabel)
                nlabel.BringToFront()
            Next
        Next

        grp.Top = grpJugador.Top + grpJugador.Height + 10
        grp.Left = grpJugador.Left
        grp.Height = grpJugador.Height
        grp.Width = grpJugador.Width
        grp.BringToFront()
        Me.Controls.Add(grp)
    End Sub
End Class
