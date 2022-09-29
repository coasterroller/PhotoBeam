Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        savepreview()
        Button3.Enabled = True
    End Sub
    Sub savepreview()
        PictureBox2.Image = New Bitmap(PictureBox1.Image, New Size(CInt(TextBox1.Text), CInt(TextBox2.Text)))
        PictureBox2.Show()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        LoadIntoArray()
    End Sub
    Sub LoadIntoArray()
        Dim myBitMap
        Dim x, y
        Dim xx = 0
        SaveFileDialog1.Filter = "Support Files (*.xml)|*.xml"
        SaveFileDialog1.ShowDialog()
        Dim file2 As New System.IO.StreamWriter(SaveFileDialog1.FileName, False)
        file2.WriteLine("<?xml version=" & """" & "1.0" & """" & " encoding=" & """" & "UTF-8" & """" & "?>")
        file2.WriteLine("<root>")
        file2.WriteLine("  <supports>")
        myBitMap = CType(PictureBox2.Image, Bitmap)
        Dim pixelColor As Color
        xx = -1
        For x = 0 To CInt(TextBox1.Text) - 1 Step CInt(TextBox3.Text)
            For y = 0 To CInt(TextBox2.Text) - 1 Step CInt(TextBox3.Text)
                xx = xx + 2
                pixelColor = myBitMap.GetPixel(x, y)
                file2.WriteLine("    <freenode id=" & """" & xx & """" & " >")
                file2.WriteLine(Replace("      <pos>" & -(CInt(TextBox1.Text) / 2) + x & " " & Random_number(CInt(TextBox5.Text), CInt(TextBox6.Text)) & " " & -(CInt(TextBox2.Text) / 2) + y & "</pos>", ",", "."))
                file2.WriteLine("    </freenode>")
                file2.WriteLine("    <footernode id=" & """" & xx + 1 & """" & " contype=" & """" & "0" & """" & " basetype=" & """" & "0" & """" & ">")
                file2.WriteLine(Replace("      <pos> " & -(CInt(TextBox1.Text) / 2) + x & " " & CInt(TextBox4.Text) & " " & -(CInt(TextBox2.Text) / 2) + y & " </pos>", ",", "."))
                file2.WriteLine("      <rotation>0</rotation>")
                file2.WriteLine("      <height_above_terrain>0.3</height_above_terrain>")
                file2.WriteLine("      <colormode_custom r=" & """" & Math.Round(pixelColor.R / 255, 2) & """" & " g=" & """" & Math.Round(pixelColor.G / 255, 2) & """" & " b=" & """" & Math.Round(pixelColor.B / 255, 2) & """" & "/>")
                file2.WriteLine("    </footernode>")
                file2.WriteLine("    <beam start=" & """" & xx + 1 & """" & " end=" & """" & xx & """" & " type=" & """" & "1" & """" & " size1=" & """" & TextBox4.Text & """" & " size2=" & """" & TextBox4.Text & """" & ">")
                file2.WriteLine("      <colormode_custom r=" & """" & Math.Round(pixelColor.R / 255, 2) & """" & " g=" & """" & Math.Round(pixelColor.G / 255, 2) & """" & " b=" & """" & Math.Round(pixelColor.B / 255, 2) & """" & "/>")
                file2.WriteLine("    </beam>")
            Next
        Next
        file2.WriteLine("  </supports>")
        file2.Write("</root>")
        file2.Close()
    End Sub
    Public Function Random_number(min_number, max_number)
        Randomize()
        Random_number = CInt(Math.Floor((max_number - min_number + 1) * Rnd())) + min_number
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Choose a Picture"
        OpenFileDialog1.Filter = "JPEG|*.jpg|PNG|*.png|Bitmap|*.bmp"
        OpenFileDialog1.ShowDialog()
        PictureBox1.ImageLocation = OpenFileDialog1.FileName.ToString
        Button2.Enabled = True
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        TextBox1.Text = 1536
        TextBox2.Text = 1536
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs)
        TextBox1.Text = 1536 / 4
        TextBox2.Text = 1536 / 4
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs)
        TextBox1.Text = 1536 / 8
        TextBox2.Text = 1536 / 8
    End Sub
    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        If TrackBar1.Value = 0 Then TrackBar1.Value = 1
        TextBox1.Text = Math.Round(1536 / TrackBar1.Value)
        TextBox2.Text = Math.Round(1536 / TrackBar1.Value)
    End Sub
    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        TextBox3.Text = Math.Round(1 + TrackBar2.Value)
    End Sub
    Private Sub TrackBar3_Scroll(sender As Object, e As EventArgs) Handles TrackBar3.Scroll
        If TrackBar3.Value = 0 Then TextBox4.Text = 0.1
        TextBox4.Text = Math.Round(0.1 + (TrackBar3.Value / 10), 2)
    End Sub
    Private Sub TrackBar4_Scroll(sender As Object, e As EventArgs) Handles TrackBar4.Scroll
        If TrackBar4.Value < 1 Then TextBox5.Text = 1
        TextBox5.Text = 1 + TrackBar4.Value
        If CInt(TextBox5.Text) > CInt(TextBox6.Text) Then TextBox5.Text = (TextBox6.Text)
    End Sub
    Private Sub TrackBar5_Scroll(sender As Object, e As EventArgs) Handles TrackBar5.Scroll
        If TrackBar5.Value < 1 Then TextBox6.Text = 1
        TextBox6.Text = 1 + TrackBar5.Value
        If CInt(TextBox6.Text) < CInt(TextBox5.Text) Then TextBox6.Text = (TextBox5.Text)
    End Sub
End Class
