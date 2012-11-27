Public Class Main
    Private Sub ButOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButOK.Click
        BarOut.Maximum = CArt.GetArtNum
        BarOut.Minimum = 1
        For i = 1 To CArt.GetArtNum
            My.Application.DoEvents()
            TextOut.Text = i
            BarOut.Value = i
            If Dir("..\App_Code\Art\" & i & ".htm") <> "" Then
                CArt.Build(i)
            End If
        Next
        TextOut.Text = "完成"
    End Sub
End Class
