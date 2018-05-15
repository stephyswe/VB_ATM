Public Class ProgrammersBtn
    Inherits Windows.Forms.Button
    Public Sub New()
        'Making Existing Button Transparent
        Me.FlatStyle = Windows.Forms.FlatStyle.Flat
        Me.FlatAppearance.BorderSize = 0
        Me.FlatAppearance.MouseDownBackColor = Color.Transparent
        Me.FlatAppearance.MouseOverBackColor = Color.Transparent
        Me.BackColor = Color.Transparent

        'ReDevelop the existing Button adding image to a transparent image
        Me.BackgroundImage = My.Resources._1button
        Me.BackgroundImageLayout = ImageLayout.Stretch
    End Sub


    Private Sub ProgrammersBtn_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        Me.BackgroundImage = My.Resources._1button_click
    End Sub


    Private Sub ProgrammersBtn_MouseHover(sender As Object, e As EventArgs) Handles Me.MouseHover
        Me.BackgroundImage = My.Resources._1button_hover
    End Sub


    Private Sub ProgrammersBtn_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        Me.BackgroundImage = My.Resources._1button
    End Sub


    Private Sub ProgrammersBtn_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        Me.BackgroundImage = My.Resources._1button
    End Sub
End Class
