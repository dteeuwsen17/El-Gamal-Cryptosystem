Public Class Form1
    Dim Darray(1000), Marray(1000), Karray(1000), Barray As String
    Function Expomod(ByVal m As Double, ByVal exp As Double, ByVal n As Double) As Double
        Dim k, i As Double
        k = m                     'ensures that vb takes the mod n of the value before the next iteration'
        For i = 1 To exp - 1         'so that the numbers are a manageable size (m^exp mod n)'
            m = (m * k) Mod n
        Next
        Return m
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Randomize() 'Ensures that randomly generated numbers do not appear in the same order everytime
        'that the user restarts the program
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim x, y, gen, p As Double
        Try
            p = TextBox3.Text
            gen = TextBox4.Text                     'Generates h=g^x mod p, c=g^y mod p, and d=g^(x*y) mod p'
            x = TextBox6.Text
            y = TextBox7.Text
            TextBox5.Text = Expomod(gen, x, p)
            TextBox8.Text = Expomod(gen, y, p)
            TextBox9.Text = Expomod(gen, x * y, p)
        Catch ex As InvalidCastException
            MessageBox.Show("A prime and random numbers must be chosen first", "Error")
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim a, b, p As Double
        Try
            p = TextBox3.Text
            b = Int((p - 2) * Rnd() + 3)        'Generates random integers for both Alice and Brad'
            TextBox7.Text = b
            a = Int((p - 2) * Rnd() + 3)
            TextBox6.Text = a
        Catch ex As InvalidCastException
            MessageBox.Show("A prime must be chosen first", "Error")
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dum As Integer
        Dim number, ch, output As String
        number = TextBox1.Text                  'converts an English message into an integer'
        TextBox2.Text = ""
        For dum = 1 To Len(number)
            ch = Mid(number, dum, 1)
            Darray(dum) = ch
            If (Asc(ch) - 86) = 0 - 54 Then
                output = ""
            Else
                output = Asc(ch) - 86
            End If
            TextBox2.Text = TextBox2.Text & output
        Next
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim isPrime As Boolean
        Dim p, b, i, j, count As Integer
        isPrime = False
        Do While isPrime = False
            count = 0
            p = 900 * Rnd() + 100
            For i = 2 To Int(Math.Sqrt(p))          'Finds a three digit prime'
                If p Mod i = 0 Then
                    Exit For
                Else
                    count = count + 1
                End If
                If count = Int(Math.Sqrt(p)) - 1 Then
                    isPrime = True
                End If
            Next
        Loop
        TextBox3.Text = p
        For i = 1 To p - 1                  'Finds the first primitive root that is greater than 4'
            For j = 1 To p - 1
                b = Expomod(i, j, p)
                If b = 1 Then
                    Exit For
                End If
            Next
            If j = p - 1 And i > 4 Then
                Exit For
            End If
        Next
        TextBox4.Text = i
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim m, s, c2, p, j, k, Mm, q As Integer
        Dim number As String
        Try
            p = TextBox3.Text
            s = TextBox9.Text
            number = TextBox2.Text
            For k = 1 To p                  'finds the inverse of "s"
                If s * k Mod p = 1 Then
                    Mm = k
                    Exit For
                End If
            Next
            TextBox10.Text = ""
            TextBox11.Text = ""
            For j = 1 To number.Length Step 2   'Since the prime number is three digits the program'
                m = Mid(number, j, 2)           'takes only two digit values at a time'
                Marray(j) = m                   'Outputs a list of cipher texts and the numerical message'
                c2 = m * s Mod p
                TextBox10.Text = TextBox10.Text & c2 & ","
                q = c2 * Mm Mod p
                TextBox11.Text = TextBox11.Text & q
            Next
        Catch ex As InvalidCastException
            MessageBox.Show("A prime and a shared secret must first be generated", "Error")
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim i, j As Integer
        Dim number, output As String
        number = TextBox11.Text
        TextBox12.Text = ""
        For i = 1 To Len(number) Step 2
            j = Mid(number, i, 2)
            Karray(i) = j
            output = Chr(j + 86)
                TextBox12.Text = TextBox12.Text & output
        Next
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        My.Computer.Audio.Play(My.Resources.missionimpossible, AudioPlayMode.Background)
    End Sub

End Class
