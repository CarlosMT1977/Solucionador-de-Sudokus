Option Explicit On
Option Strict On


Public Class Form1
    Private MatrizDeSudoku(80) As Integer
    Private MatrizDeFijos(80) As Boolean
    Private CasillaActual As Integer
    Private SudokuResuelto As Boolean = False
    Private SudokuSinSolucion As Boolean = False
    Private MatrizDeTextBoxes(80) As TextBox

    Const DistanciaIzquierda As Integer = 50
    Const DistanciaArriba As Integer = 50
    Const DistanciaMinima As Integer = 10
    Const DistanciaMaxima As Integer = 30

    Private Function QueFilaEs(NumeroDeCasilla As Integer) As Integer
        If NumeroDeCasilla < 0 Or NumeroDeCasilla > 80 Then CortamosConError("El número de casilla tiene que ser un número entre 0 y 80, ambos inclusive", "Número de casilla inválido") : Stop : End
        Dim Resultado As Integer = NumeroDeCasilla \ 9
        If Resultado < 0 Or Resultado > 8 Then CortamosConError("El resultado tiene que ser un número entre 0 y 8, ambos inclusive", "Resultado inválido") : Stop : End
        Return Resultado
    End Function

    Private Function QueColumnaEs(NumeroDeCasilla As Integer) As Integer
        If NumeroDeCasilla < 0 Or NumeroDeCasilla > 80 Then CortamosConError("El número de casilla tiene que ser un número entre 0 y 80, ambos inclusive", "Número de casilla inválido") : Stop : End
        Dim Resultado As Integer = NumeroDeCasilla Mod 9
        If Resultado < 0 Or Resultado > 8 Then CortamosConError("El resultado tiene que ser un número entre 0 y 8, ambos inclusive", "Resultado inválido") : Stop : End
        Return Resultado
    End Function

    Private Function QueTresPorTresEs(NumeroDeCasilla As Integer) As Integer
        If NumeroDeCasilla < 0 Or NumeroDeCasilla > 80 Then CortamosConError("El número de casilla tiene que ser un número entre 0 y 80, ambos inclusive", "Número de casilla inválido") : Stop : End

        Dim Resultado As Integer = 3 * (NumeroDeCasilla \ 27) + ((NumeroDeCasilla Mod 9) \ 3)

        If Resultado < 0 Or Resultado > 8 Then CortamosConError("El resultado tiene que ser un número entre 0 y 8, ambos inclusive", "Resultado inválido") : Stop : End
        Return Resultado
    End Function


    Private Function EsFactibleLaFila(NumeroDeFila As Integer) As Boolean
        If NumeroDeFila < 0 Or NumeroDeFila > 8 Then CortamosConError(0) : Stop : End
        Dim NumeroDeAparicionesDeCada(9) As Integer : InicializarMatriz(NumeroDeAparicionesDeCada)
        Dim Contador, Inicio, Final As Integer
        Inicio = 9 * NumeroDeFila
        Final = 9 * NumeroDeFila + 8
        For Contador = Inicio To Final
            If MatrizDeSudoku(Contador) <> 0 Then
                NumeroDeAparicionesDeCada(MatrizDeSudoku(Contador)) += 1
                If NumeroDeAparicionesDeCada(MatrizDeSudoku(Contador)) > 1 Then Return False
            End If
        Next
        If Contador <> 9 * (NumeroDeFila + 1) Then CortamosConError(0) : Stop : End
        Return True
    End Function

    Private Function EsFactibleLaColumna(NumeroDeColumna As Integer) As Boolean
        If NumeroDeColumna < 0 Or NumeroDeColumna > 8 Then CortamosConError(0) : Stop : End
        Dim NumeroDeAparicionesDeCada(9) As Integer : InicializarMatriz(NumeroDeAparicionesDeCada)
        Dim Contador, Inicio, Final As Integer
        Inicio = NumeroDeColumna
        Final = 72 + NumeroDeColumna
        For Contador = Inicio To Final Step 9
            If MatrizDeSudoku(Contador) <> 0 Then
                NumeroDeAparicionesDeCada(MatrizDeSudoku(Contador)) += 1
                If NumeroDeAparicionesDeCada(MatrizDeSudoku(Contador)) > 1 Then Return False
            End If
        Next
        If Contador <> NumeroDeColumna + 81 Then CortamosConError(0) : Stop : End
        Return True
    End Function

    Private Function EsFactibleElTresPorTres(NumeroDeTresPorTres As Integer) As Boolean
        If NumeroDeTresPorTres < 0 Or NumeroDeTresPorTres > 8 Then CortamosConError(0) : Stop : End
        Dim NumeroDeAparicionesDeCada(9) As Integer : InicializarMatriz(NumeroDeAparicionesDeCada)
        Dim NumeroDeCasilla, CuentaFilas, CuentaColumnas As Integer
        For CuentaFilas = 3 * (NumeroDeTresPorTres \ 3) To 3 * (NumeroDeTresPorTres \ 3) + 2
            For CuentaColumnas = 3 * (NumeroDeTresPorTres Mod 3) To 3 * (NumeroDeTresPorTres Mod 3) + 2
                If MatrizDeSudoku(QueCasillaEs(CuentaFilas, CuentaColumnas)) <> 0 Then
                    NumeroDeAparicionesDeCada(MatrizDeSudoku(QueCasillaEs(CuentaFilas, CuentaColumnas))) += 1
                    If NumeroDeAparicionesDeCada(MatrizDeSudoku(QueCasillaEs(CuentaFilas, CuentaColumnas))) > 1 Then Return False
                End If
            Next
        Next
        If CuentaFilas <> 3 * (NumeroDeTresPorTres \ 3) + 3 Then CortamosConError(0) : Stop : End
        If CuentaColumnas <> 3 * (NumeroDeTresPorTres Mod 3) + 3 Then CortamosConError(0) : Stop : End
        Return True
        End Function

    Private Function QueCasillaEs(NumeroDeFila As Integer, NumeroDeColumna As Integer) As Integer
        If NumeroDeFila < 0 Or NumeroDeFila > 8 Or NumeroDeColumna < 0 Or NumeroDeColumna > 8 Then CortamosConError(0) : Stop : End
        Dim Resultado As Integer
        Resultado = 9 * NumeroDeFila + NumeroDeColumna
        If Resultado < 0 Or Resultado > 80 Then CortamosConError(0) : Stop : End
        Return Resultado
    End Function

    Private Function EsFactibleLaCasilla(NumeroDeCasilla As Integer) As Boolean
        If NumeroDeCasilla < 0 Or NumeroDeCasilla > 80 Then CortamosConError(0) : Stop : End
        If Not EsFactibleLaFila(QueFilaEs(NumeroDeCasilla)) Then Return False
        If Not EsFactibleLaColumna(QueColumnaEs(NumeroDeCasilla)) Then Return False
        If Not EsFactibleElTresPorTres(QueTresPorTresEs(NumeroDeCasilla)) Then Return False
        Return True
    End Function


    Private Sub Avanzamos()
        CasillaActual += 1
        If CasillaActual > 80 Then Exit Sub
        If Not MatrizDeFijos(CasillaActual) Then MatrizDeSudoku(CasillaActual) = 1
    End Sub

    Private Sub Retrocedemos()
        Do While MatrizDeFijos(CasillaActual) OrElse MatrizDeSudoku(CasillaActual) = 9
            If Not MatrizDeFijos(CasillaActual) Then MatrizDeSudoku(CasillaActual) = 0
            CasillaActual -= 1
            If CasillaActual < 0 Then Exit Do
        Loop
        If CasillaActual < 0 Then Exit Sub
        MatrizDeSudoku(CasillaActual) += 1
    End Sub


    Private Sub ResolverSudoku()
        SudokuResuelto = False
        SudokuSinSolucion = False
        Dim Contador As Integer
        For Contador = 0 To 80
            If MatrizDeSudoku(Contador) < 0 Or MatrizDeSudoku(Contador) > 9 Then SudokuSinSolucion = True : Exit Sub
            If MatrizDeSudoku(Contador) = 0 Then MatrizDeFijos(Contador) = False Else MatrizDeFijos(Contador) = True
            If MatrizDeFijos(Contador) And Not EsFactibleLaCasilla(Contador) Then SudokuSinSolucion = True : Exit Sub
        Next
        CasillaActual = 0
        If Not MatrizDeFijos(CasillaActual) Then MatrizDeSudoku(CasillaActual) = 1
        Do Until CasillaActual < 0 Or CasillaActual > 80
            If EsFactibleLaCasilla(CasillaActual) Then
                Avanzamos()
            Else
                Retrocedemos()
            End If
        Loop
        If CasillaActual = -1 Then
            SudokuSinSolucion = True
        ElseIf CasillaActual = 81 Then
            SudokuResuelto = True
        Else
            CortamosConError(0) : Stop : End
        End If
        If SudokuResuelto And SudokuSinSolucion Then CortamosConError(0) : Stop : End
    End Sub



    Private Sub CortamosConError(MensajeDeError As String, Optional TituloDeMensaje As String = vbNullString)
        MessageBox.Show(MensajeDeError, TituloDeMensaje, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub CortamosConError(NumeroDeError As Integer)
        Select Case NumeroDeError
            Case 0 : CortamosConError("Si estamos aquí, es que algo ha salido mal", "No deberíamos estar aquí")
        End Select
    End Sub

    Private Sub AdvertimosAlUsuario(MensajeDeAdvertencia As String, Optional TituloDeMensaje As String = vbNullString)
        MessageBox.Show(MensajeDeAdvertencia, TituloDeMensaje, MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub InicializarMatriz(ByRef Matriz() As Integer)
        Dim Contador As Integer
        For Contador = 0 To Matriz.GetUpperBound(0)
            Matriz(Contador) = 0
        Next
    End Sub

    Private Sub InicializarMatriz(ByRef Matriz() As Boolean)
        Dim Contador As Integer
        For Contador = 0 To Matriz.GetUpperBound(0)
            Matriz(Contador) = False
        Next
    End Sub

    Private Function ImpresionDeMatriz(Matriz() As Integer) As String
        Dim Contador As Integer
        Dim CadenaDeTexto As String = vbNullString
        For Contador = 0 To Matriz.GetUpperBound(0)
            CadenaDeTexto &= Matriz(Contador)
            If (Contador + 1) Mod 9 = 0 Then
                CadenaDeTexto &= vbCrLf
            ElseIf (Contador + 1) Mod 3 = 0 Then
                CadenaDeTexto &= "   "
            End If
            If (Contador + 1) Mod 27 = 0 Then CadenaDeTexto &= vbCrLf
        Next
        Return CadenaDeTexto
    End Function

    Private Function ImpresionDeMatriz(Matriz() As Boolean) As String
        ' Imprime 1 para True y 0 para False    
        Dim Contador As Integer
        Dim CadenaDeTexto As String = vbNullString
        For Contador = 0 To Matriz.GetUpperBound(0)
            If Matriz(Contador) = True Then CadenaDeTexto &= "1" Else CadenaDeTexto &= "0"
            If (Contador + 1) Mod 9 = 0 Then
                CadenaDeTexto &= vbCrLf
            ElseIf (Contador + 1) Mod 3 = 0 Then
                CadenaDeTexto &= "   "
            End If
            If (Contador + 1) Mod 27 = 0 Then CadenaDeTexto &= vbCrLf
        Next
        Return CadenaDeTexto

    End Function


    Private Sub AlEntrarEnElCuadroDeTexto(sender As Object, e As EventArgs)
        Dim ControlEnlazado As TextBox = CType(sender, TextBox)
        ControlEnlazado.SelectAll()
    End Sub

    Private Sub AlSalirDeUnCuadroDeTexto(sender As Object, e As EventArgs)
        Dim ControlEnlazado As TextBox = CType(sender, TextBox)
        If ControlEnlazado.Text = " " Then ControlEnlazado.Text = vbNullString
        If ControlEnlazado.Text <> vbNullString And (ControlEnlazado.Text < "1" Or ControlEnlazado.Text > "9") Then
            AdvertimosAlUsuario("Sólo puedes poner un número entero del 1 al 9, ambos inclusive", "Contenido inválido")
            ControlEnlazado.Focus()
        End If
    End Sub


    Private Sub CrearLaMatrizDeControles()
        Dim XAux, YAux As Integer
        Dim Contador As Integer
        InicializarMatriz(MatrizDeSudoku)
        InicializarMatriz(MatrizDeFijos)
        For Contador = 0 To 80
            MatrizDeTextBoxes(Contador) = New TextBox
            MatrizDeTextBoxes(Contador).MaxLength = 1
            MatrizDeTextBoxes(Contador).BackColor = Color.White
            MatrizDeTextBoxes(Contador).TextAlign = HorizontalAlignment.Center
            MatrizDeTextBoxes(Contador).Size = New Size(MatrizDeTextBoxes(Contador).Size.Height, MatrizDeTextBoxes(Contador).Size.Height)
            XAux = DistanciaIzquierda + (Contador Mod 9) * (MatrizDeTextBoxes(Contador).Size.Width + DistanciaMinima) + ((Contador Mod 9) \ 3) * (DistanciaMaxima - DistanciaMinima)
            YAux = DistanciaArriba + (Contador \ 9) * (MatrizDeTextBoxes(Contador).Size.Height + DistanciaMinima) + (Contador \ 27) * (DistanciaMaxima - DistanciaMinima)
            MatrizDeTextBoxes(Contador).Location = New Point(XAux, YAux)
            MatrizDeTextBoxes(Contador).Name = CType(Contador, String)
            Me.Controls.Add(MatrizDeTextBoxes(Contador))

            AddHandler MatrizDeTextBoxes(Contador).Enter, AddressOf AlEntrarEnElCuadroDeTexto
            AddHandler MatrizDeTextBoxes(Contador).Click, AddressOf AlEntrarEnElCuadroDeTexto

            AddHandler MatrizDeTextBoxes(Contador).Leave, AddressOf AlSalirDeUnCuadroDeTexto

        Next

        XAux = 2 * DistanciaIzquierda + 9 * MatrizDeTextBoxes(0).Width + 6 * DistanciaMinima + 2 * DistanciaMaxima
        YAux = 2 * DistanciaArriba + 9 * MatrizDeTextBoxes(0).Height + 6 * DistanciaMinima + 3 * DistanciaMaxima + Button1.Size.Height
        Me.ClientSize = New Size(XAux, YAux)
        XAux = Me.ClientSize.Width \ 2 - Button1.Size.Width \ 2
        YAux = Me.ClientSize.Height - DistanciaArriba - Button1.Size.Height
        Button1.Location = New Point(XAux, YAux)
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Contador As Integer
        InicializarMatriz(MatrizDeSudoku)
        InicializarMatriz(MatrizDeFijos)
        If Button1.Text = "Resetear cuadro" Then
            For Contador = 0 To 80
                MatrizDeTextBoxes(Contador).BackColor = Color.White
                MatrizDeTextBoxes(Contador).Text = vbNullString
            Next
            Button1.Text = "Resolver Sudoku"
            Exit Sub
        End If


        For Contador = 0 To 80
            If MatrizDeTextBoxes(Contador).Text <> vbNullString Then
                MatrizDeFijos(Contador) = True
                MatrizDeSudoku(Contador) = CType(MatrizDeTextBoxes(Contador).Text, Integer)
            End If
        Next

        ResolverSudoku()
        If SudokuResuelto And SudokuSinSolucion Then CortamosConError(0) : Stop : End
        If SudokuSinSolucion Then
            AdvertimosAlUsuario("Este Sudoku no tiene solución")
        ElseIf SudokuResuelto Then
            For Contador = 0 To 80
                If Not MatrizDeFijos(Contador) Then MatrizDeTextBoxes(Contador).BackColor = Color.Aqua
                MatrizDeTextBoxes(Contador).Text = CType(MatrizDeSudoku(Contador), String)
            Next
            Button1.Text = "Resetear cuadro"
        Else
            CortamosConError(0) : Stop : End
        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CrearLaMatrizDeControles()

    End Sub
End Class
