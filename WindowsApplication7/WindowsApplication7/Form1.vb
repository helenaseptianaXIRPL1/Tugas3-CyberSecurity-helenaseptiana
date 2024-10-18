Imports System.IO
Imports System.Text

Public Class Notepad
    Private Sub Notepad_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Me.Text = "Notepad"
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click
        TextBox1.Clear()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripMenuItem.Click
        Using openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "All Files (*.*)|*.*"
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Try
                    Dim fileContent As String = System.IO.File.ReadAllText(openFileDialog.FileName)
                    TextBox1.Text = fileContent ' Menampilkan isi file
                Catch ex As Exception
                    MessageBox.Show("Error reading file: " & ex.Message, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub OpenFolderToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OpenFolderToolStripMenuItem.Click
        Using folderBrowserDialog As New FolderBrowserDialog()
            If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
                Dim files As String() = Directory.GetFiles(folderBrowserDialog.SelectedPath)
                ListBox1.Items.Clear() ' Mengosongkan ListBox sebelumnya

                For Each file In files
                    ListBox1.Items.Add(Path.GetFileName(file)) ' Menambahkan nama file ke ListBox
                Next

                If ListBox1.Items.Count > 0 Then
                    MessageBox.Show("Files found in the selected folder. Select one to open.", "Files Loaded", MessageBoxButtons.OK)
                Else
                    MessageBox.Show("No files found in the selected folder.", "No Files", MessageBoxButtons.OK)
                End If
            End If
        End Using
    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedItem IsNot Nothing Then
            Dim selectedFile As String = ListBox1.SelectedItem.ToString()
            Dim folderPath As String = Path.GetDirectoryName(ListBox1.Items(0).ToString())
            Dim filePath As String = Path.Combine(folderPath, selectedFile)

            Try
                Dim fileContent As String = System.IO.File.ReadAllText(filePath)
                TextBox1.Text = fileContent ' Menampilkan isi file
            Catch ex As Exception
                MessageBox.Show("Unable to display content of the file: " & ex.Message, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveToolStripMenuItem.Click
        Using saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim encryptedText As String = Encrypt(TextBox1.Text)
                System.IO.File.WriteAllText(saveFileDialog.FileName, encryptedText)
            End If
        End Using
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    ' Enkripsi menggunakan nilai ASCII
    Private Function Encrypt(ByVal input As String) As String
        Dim key As Integer = 3 ' Kunci pergeseran
        Dim output As New StringBuilder()

        For Each c As Char In input
            ' Menggeser nilai ASCII setiap karakter
            output.Append(Chr(Asc(c) + key))
        Next

        Return output.ToString()
    End Function
End Class
