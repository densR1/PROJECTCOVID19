Public Class Form1
    Dim sqlnya As String
    Sub panggildata()
        Konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM COVID", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "COVID")
        DataGridView1.DataSource = DS.Tables("COVID")
        DataGridView1.Enabled = True
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggildata()

    End Sub
    Sub jalan()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call Konek()
        objcmd.Connection = conn
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnya
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        Nis.Text = ""
        Nama.Text = ""
        Umur.Text = ""
        Pekerjaan.Text = ""
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button1.Click
        Dim Jumlah As Integer
        Jumlah = 0
        If CheckBox1.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox3.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox5.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox7.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox9.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox11.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox13.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox15.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox17.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox19.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox21.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox23.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox25.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox27.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox29.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox31.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox33.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox35.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox37.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox39.Checked = True Then
            Jumlah = Jumlah + 1
        End If
        If CheckBox41.Checked = True Then
            Jumlah = Jumlah + 1
        End If

        sqlnya = "insert into COVID (Nis, Nama, Umur, Pekerjaan, jumlah)values('" & Nis.Text & "','" & Nama.Text & "','" & Umur.Text & "','" & Pekerjaan.Text & "','" & Jumlah & "')"
        Call jalan()
        MsgBox("Data Tersimpan")
        Call panggildata()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        sqlnya = "delete from COVID (Nis, Nama, Umur, Pekerjaan, jumlah)values('" & Nis.Text & "','" & Nama.Text & "','" & Umur.Text & "','" & Pekerjaan.Text & "','" & Jumlah & "')"
        Call jalan()
        MsgBox("Data Terhapus")
        Call panggildata()
    End Sub

    Private Function Jumlah() As String

    End Function

End Class
