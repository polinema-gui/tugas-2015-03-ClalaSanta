﻿Public Class Form1

    Private Sub ProblemBintangToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProblemBintangToolStripMenuItem.Click
        Form2.Show()
    End Sub

    Private Sub MatriksKalkulatorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MatriksKalkulatorToolStripMenuItem.Click
        Form3.Show()
    End Sub

    Private Sub KeluarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KeluarToolStripMenuItem.Click
        Dim pilih = MessageBox.Show("Are you sure to quit ?", "Konfirmasi", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If pilih = Windows.Forms.DialogResult.OK Then
            Environment.Exit(0)
        Else

        End If
    End Sub
    Private Sub AplikasiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AplikasiToolStripMenuItem.Click



    End Sub
End Class
