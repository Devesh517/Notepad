Public Class Form1
    Dim filestate As Boolean
    Dim curfilename As String
    Dim cutid As Boolean = False
    Private isDarkMode As Boolean = False
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterParent
        Me.StartPosition = FormStartPosition.CenterScreen
        RichTextBox1.ScrollBars = RichTextBoxScrollBars.Both
        RichTextBox1.WordWrap = False
        Label1.Text = "Date: " & DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt")
        Timer1.Start()
         End Sub
    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub
    Private Sub HScrollBar1_Scroll(sender As Object, e As ScrollEventArgs)

    End Sub
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        OpenFileDialog1.InitialDirectory = "d:\"
        OpenFileDialog1.Filter = "TextFile|*.txt|Document|*.doc|All Files|*.*"
        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            RichTextBox1.LoadFile(OpenFileDialog1.FileName)
            curfilename = OpenFileDialog1.FileName
            filestate = False
        End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDialog1.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If filestate = True Then
            RichTextBox1.SaveFile(curfilename)
        Else
            SaveFileDialog1.InitialDirectory = "d:\"
            SaveFileDialog1.Filter = "Text File|*.txt|Document|*.doc|All Files|*.*"
            If SaveFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
                RichTextBox1.SaveFile(SaveFileDialog1.FileName)
                curfilename = SaveFileDialog1.FileName
                filestate = True
            End If
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveFileDialog1.InitialDirectory = "d:\"
        SaveFileDialog1.Filter = "Text File|*.txt|Document|*.doc|All Files|*.*"
        If SaveFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            RichTextBox1.SaveFile(SaveFileDialog1.FileName)
            curfilename = SaveFileDialog1.FileName
            filestate = True
        End If
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        If filestate = True Then
            Dim x As Integer
            x = MsgBox("Do you want to Save file.. ", MsgBoxStyle.YesNoCancel, "My Project")
            If x = 6 Then
                RichTextBox1.SaveFile(curfilename)
                RichTextBox1.Clear()
                filestate = False
                RichTextBox1.Focus()
            ElseIf x = 7 Then
                RichTextBox1.Clear()
                curfilename = ""
                filestate = False
                RichTextBox1.Focus()
            End If
        End If

    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        RichTextBox1.Refresh()
    End Sub

    Private Sub FontColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionColor = ColorDialog1.Color
    End Sub

    Private Sub FontStyleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontStyleToolStripMenuItem.Click
        FontDialog1.ShowDialog()
        RichTextBox1.SelectionFont = FontDialog1.Font
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionColor = ColorDialog1.Color
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        If RichTextBox1.SelectionFont IsNot Nothing Then
            Dim currentFont As Font = RichTextBox1.SelectionFont
            Dim newFontStyle As FontStyle = currentFont.Style Xor FontStyle.Bold
            RichTextBox1.SelectionFont = New Font(currentFont, newFontStyle)
        End If
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If RichTextBox1.SelectionFont IsNot Nothing Then
            Dim currentFont As Font = RichTextBox1.SelectionFont
            Dim newFontStyle As FontStyle = currentFont.Style Xor FontStyle.Underline
            RichTextBox1.SelectionFont = New Font(currentFont, newFontStyle)
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If RichTextBox1.SelectionFont IsNot Nothing Then
            Dim currentFont As Font = RichTextBox1.SelectionFont
            Dim newFontStyle As FontStyle = currentFont.Style Xor FontStyle.Italic
            RichTextBox1.SelectionFont = New Font(currentFont, newFontStyle)
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label1.Text = DateTime.Now.ToString("hh:mm:ss tt")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If isDarkMode = False Then
            Me.BackColor = Color.FromArgb(30, 30, 30)
            Button1.BackColor = Color.DimGray
            Button1.ForeColor = Color.White
            MenuStrip1.BackColor = Color.Black
            MenuStrip1.ForeColor = Color.White
            ToolStrip1.BackColor = Color.DimGray
            ToolStrip1.ForeColor = Color.White
            Label1.ForeColor = Color.White
            RichTextBox1.BackColor = Color.Black
            RichTextBox1.ForeColor = Color.White
            isDarkMode = True
        Else
            Me.BackColor = Color.White
            Button1.BackColor = Color.LightGray
            Button1.ForeColor = Color.Black
            Label1.ForeColor = Color.Black
            MenuStrip1.BackColor = SystemColors.Control
            MenuStrip1.ForeColor = Color.Black
            ToolStrip1.BackColor = SystemColors.Control
            ToolStrip1.ForeColor = Color.Black
            RichTextBox1.BackColor = Color.White
            RichTextBox1.ForeColor = Color.Black
            isDarkMode = False
        End If
    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub MenuStrip1_ItemClicked_1(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class
