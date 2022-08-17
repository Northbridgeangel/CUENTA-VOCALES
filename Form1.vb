Public Class Form1

    ' CUENTAVOCALES 1º OPCIÓN

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '--------------------------------------------------------------------------------
        '------------ AHORA LIMPIO LOS TEXTBOX -------------------------------------------
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""

        '----------- DIMENSIONO LAS VARIABLES ------------------------------------------

        Dim FRASE As String
        Dim TAMAÑO As Integer
        Dim LETRA As Char
        Dim CT, CA, CE, CI, CO, CU As Integer

        '------------ PIDO FRASE Y LA MIDO ---------------------------------------------

        FRASE = InputBox("Mete la Frase o Palabra a analizar", "AÑADE FRASE O PALABRA")
        '---------- UCASE CONVIERTE LA VARIABLE (FRASE) EN MAYUSCULAS --------------------
        FRASE = UCase(FRASE)
        '------------ LCASE CONVIERTE LA VARIABLE A MINÚSCULAS ---------------------------
        TAMAÑO = Len(FRASE)


        '---------- VOY SACANDO LETRA A LETRA DESE 1 HASTA LO QUE MIDE -------------------

        For I = 1 To TAMAÑO
            LETRA = Mid(FRASE, I, 1)
            If LETRA = "A" Or LETRA = "E" Or LETRA = "I" Or LETRA = "O" Or LETRA = "U" Then
                '----------- CONTADOR TOTAL DE VOCALES (CT) = CA + CE + CI + CO + CU ------
                CT = CT + 1
            End If
            If LETRA = "A" Then
                CA = CA + 1
            End If
            If LETRA = "E" Then
                CE = CE + 1
            End If
            If LETRA = "I" Then
                CI = CI + 1
            End If
            If LETRA = "O" Then
                CO = CO + 1
            End If
            If LETRA = "U" Then
                CU = CU + 1
            End If
        Next

        '------------------ MUESTRO EN LA CAJA DE TEXTO LOS CONTADORES ------------------------

        TextBox1.text = CT
        TextBox2.text = CA
        TextBox3.text = CE
        TextBox4.text = CI
        TextBox5.text = CO
        TextBox6.text = CU

    End Sub




    ' CUENTAVOCALES 2º OPCIÓN

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        '--------------------------------------------------------------------------------
        '------------ AHORA LIMPIO LOS TEXTBOX -------------------------------------------
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""


        '----------- DIMENSIONO LAS VARIABLES ------------------------------------------

        Dim FRASE As String
        Dim TAMAÑO As Integer
        Dim LETRA As Char
        Dim CT, CA, CE, CI, CO, CU As Integer

        '------------ PIDO FRASE Y LA MIDO ---------------------------------------------

        FRASE = InputBox("Mete la Frase o Palabra a analizar", "AÑADE FRASE O PALABRA")
        '---------- UCASE CONVIERTE LA VARIABLE (FRASE) EN MAYUSCULAS --------------------
        FRASE = UCase(FRASE)
        '------------ LCASE CONVIERTE LA VARIABLE A MINÚSCULAS ---------------------------
        TAMAÑO = Len(FRASE)


        '---------- VOY SACANDO LETRA A LETRA DESE 1 HASTA LO QUE MIDE -------------------

        For I = 1 To TAMAÑO
            LETRA = Mid(FRASE, I, 1)

            '------------ SELECT CASE COMO CONTADOR DE VOCALES -------------------------------

            Select Case LETRA
                Case "A"
                    CA = CA + 1
                Case "E"
                    CE = CE + 1
                Case "I"
                    CI = CI + 1
                Case "O"
                    CO = CO + 1
                Case "U"
                    CU = CU + 1
            End Select
        Next
        CT = CA + CE + CI + CO + CU


        '------------------ MUESTRO EN LA CAJA DE TEXTO LOS CONTADORES ------------------------

        TextBox1.Text = CT
        TextBox2.Text = CA
        TextBox3.Text = CE
        TextBox4.Text = CI
        TextBox5.Text = CO
        TextBox6.Text = CU

    End Sub



    ' FUNCIONES DE CADENA

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox8.Select()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim palabra As String
        palabra = TextBox8.Text
        '------------------------------------------- LONGITUD
        TextBox9.Text = Len(palabra)
        '------------------------------------------- 1º CARACTER
        TextBox10.Text = Mid(palabra, 1, 1)
        '------------------------------------------- ULTIMO
        TextBox11.Text = Mid(palabra, Len(palabra), 1)
        '------------------------------------------- 2 AL 6 CARACTER
        TextBox12.Text = Mid(palabra, 2, 6)
        '------------------------------------------- DONDE ESTA LA PRIMERA A
        TextBox13.Text = InStr(UCase(palabra), "A")
        '------------------------------------------- MINÚSCULAS
        TextBox14.Text = LCase(palabra)
        '------------------------------------------- MAYÚSCULAS
        TextBox15.Text = UCase(palabra)
        '------------------------------------------- 1º EN MAYÚSCULAS
        TextBox16.Text = UCase(Mid(palabra, 1, 1))
    End Sub

    Private Sub palabra_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    'MEDIR ESPACIOS

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '------------ PIDO FRASE Y LA MIDO ---------------------------------------------
        'Dim palabra As String
        'palabra = InputBox("Mete la Frase o Palabra a analizar", "AÑADE FRASE O PALABRA")
        'CONESPACIO.Text = palabra
        CONESPACIO.Text = InputBox("Mete la Frase o Palabra a analizar", "AÑADE FRASE O PALABRA")
        ResultadoCE.Text = Len(CONESPACIO.Text)
        SINESPACIO.Text = Trim(CONESPACIO.Text)
        ResultadoSE.Text = Len(SINESPACIO.Text)

    End Sub

    Private Sub CONESPACIO_TextChanged(sender As Object, e As EventArgs) Handles CONESPACIO.TextChanged

    End Sub

    Private Sub SINESPACIO_TextChanged(sender As Object, e As EventArgs) Handles SINESPACIO.TextChanged

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim OPERADOR1, RESTO As Double
        OPERADOR1 = InputBox("METE EL Nº OBJETO DEL ANÁLISIS")

        If OPERADOR1 Mod 2 = 0 Then
            MsgBox("¡¡¡PAR!!!")
        Else
            MsgBox("¡¡¡IMPAR!!!")
        End If
    End Sub
End Class
