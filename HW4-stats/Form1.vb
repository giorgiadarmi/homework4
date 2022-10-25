Public Class Form1
    Public b As Bitmap
    Public g As Graphics
    Public r As New Random
    Public PenAbsolute As New Pen(Color.OrangeRed, 2)
    Public PenRelative As New Pen(Color.Green, 2)
    Public PenNormalized As New Pen(Color.Blue, 2)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.b = New Bitmap(Me.PictureBox1.Width, Me.PictureBox1.Height)
        Me.g = Graphics.FromImage(b)
        Me.g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        Me.g.Clear(Color.White)

        Dim TrialsCount As Integer = 100
        Dim NumerOfTrajectories As Integer = 30
        Dim SuccessProbability As Double = 0.5

        Dim minX As Double = 0
        Dim maxX As Double = TrialsCount
        Dim minY As Double = 0
        Dim maxY As Double = TrialsCount

        Dim FreqAbs As Integer = 0

        Dim VirtualWindow As New Rectangle(50, 50, Me.b.Width - 40, Me.b.Height - 40)

        g.DrawRectangle(Pens.DarkSlateGray, VirtualWindow)

        For i As Integer = 1 To NumerOfTrajectories

            Dim PuntiA As New List(Of Point)
            Dim PuntiR As New List(Of Point)
            Dim PuntiN As New List(Of Point)

            Dim Succ As Double = 0
            Dim NotSucc As Double = 0
            Dim Trials As Integer = 0

            For X As Integer = 1 To TrialsCount
                Dim Uniform As Double = r.NextDouble
                If Uniform < SuccessProbability Then
                    Succ = Succ + 1
                    FreqAbs = FreqAbs + 1
                    Trials = Trials + 1
                Else
                    NotSucc = NotSucc + 1
                    Trials = Trials + 1
                End If
                Dim xDevice As Integer = FromXRealToXVirtual(X, minX, maxX, VirtualWindow.Left, VirtualWindow.Width)
                Dim YDevice As Integer = FromYRealToYVirtual(Succ, minY, maxY, VirtualWindow.Top, VirtualWindow.Height)
                PuntiA.Add(New Point(xDevice, YDevice))

                Dim Relative As Integer = Succ * TrialsCount / (X + 1)
                Dim YDeviceR As Integer = FromYRealToYVirtual(Relative, minY, maxY, VirtualWindow.Top, VirtualWindow.Height)
                PuntiR.Add(New Point(xDevice, YDeviceR))

                Dim Normalized As Double = Succ * (Math.Sqrt(TrialsCount)) / Math.Sqrt(X + 1)
                Dim YDeviceN As Integer = FromYRealToYVirtual(Normalized, minY, maxY * SuccessProbability, VirtualWindow.Top, VirtualWindow.Height)
                PuntiN.Add(New Point(xDevice, YDeviceN))
            Next
            g.DrawLines(PenAbsolute, PuntiA.ToArray)
            g.DrawLines(PenRelative, PuntiR.ToArray)
            g.DrawLines(PenNormalized, PuntiN.ToArray)
        Next

        Me.PictureBox1.Image = b

    End Sub

    Function FromXRealToXVirtual(X As Double,
                                 minX As Double, maxX As Double,
                                 Left As Integer, W As Integer) As Integer

        If (maxX - minX) = 0 Then
            Return 0
        End If

        Return Left + W * (X - minX) / (maxX - minX)

    End Function

    Function FromYRealToYVirtual(Y As Double,
                                minY As Double, maxY As Double,
                                Top As Integer, H As Integer) As Integer

        If (maxY - minY) = 0 Then
            Return 0
        End If

        Return Top + H - H * (Y - minY) / (maxY - minY)

    End Function

End Class