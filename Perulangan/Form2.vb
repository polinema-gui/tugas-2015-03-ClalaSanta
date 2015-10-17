Public Class Form2

    Dim kolom, Bawah, Atas, jumlahKolom, jumlahBaris As Integer
    Dim hasil As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        hasil = ""
        jumlahBaris = Convert.ToInt32(TextBox1.Text)
        jumlahKolom = (jumlahBaris * 2) - 1
        If ComboBox1.Text = "Piramyd" Then
            Atas = jumlahBaris + 1
            Bawah = jumlahBaris - 1
            For baris As Integer = 1 To jumlahBaris
                kolom = 1
                For kolom As Integer = 1 To jumlahKolom
                    If kolom > Bawah And kolom < Atas Then
                        hasil &= "**"
                    Else
                        hasil &= " "
                    End If
                Next
                Bawah -= 1
                Atas += 1
                hasil &= vbCrLf
            Next

        ElseIf ComboBox1.Text = "Hollow Piramyd" Then
            Atas = jumlahBaris
            Bawah = jumlahBaris
            For baris As Integer = 1 To jumlahBaris
                kolom = 1
                If baris = jumlahBaris Then
                    For kolom As Integer = 1 To jumlahKolom
                        hasil &= "*"
                    Next
                Else
                    For kolom As Integer = 1 To jumlahKolom
                        If kolom = Bawah Or kolom = Atas Then
                            hasil &= "*"
                        Else
                            hasil &= " "
                        End If
                    Next
                End If
                Bawah -= 1
                Atas += 1
                hasil &= vbCrLf
            Next

        ElseIf ComboBox1.Text = "Inverted Piramyd" Then
            Atas = jumlahKolom
            Bawah = 1
            For baris As Integer = 1 To jumlahBaris
                kolom = 1
                For kolom As Integer = 1 To jumlahKolom
                    If kolom >= Bawah And kolom <= Atas Then
                        hasil &= "*"
                    Else
                        hasil &= " "
                    End If
                Next
                Bawah += 1
                Atas -= 1
                hasil &= vbCrLf
            Next
        ElseIf ComboBox1.Text = "Hollow Inverted Piramyd" Then
            Atas = jumlahKolom
            Bawah = 1
            For baris As Integer = 1 To jumlahBaris
                kolom = 1
                If baris = 1 Then
                    For kolom As Integer = 1 To jumlahKolom
                        hasil &= "*"
                    Next
                Else
                    For kolom As Integer = 1 To jumlahKolom
                        If kolom = Bawah Or kolom = Atas Then
                            hasil &= "*"
                        Else
                            hasil &= " "
                        End If
                    Next
                End If
                Bawah += 1
                Atas -= 1
                hasil &= vbCrLf
            Next
        End If
        TextBox2.Text = hasil
    End Sub
End Class