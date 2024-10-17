Imports System.IO

Class MainWindow
    Private Sub Button1_Click(sender As Object, e As RoutedEventArgs)
        Dim largestFolder As String = ""
        Dim largestSize As Long = 0

        ' Specify the drive or directory to search
        Dim rootPath As String = "C:\"

        Try
            ' Get all directories in the specified path
            Dim directories As String() = Directory.GetDirectories(rootPath, "*", SearchOption.AllDirectories)

            For Each dir As String In directories
                Dim folderSize As Long = GetDirectorySize(dir)

                ' Check if this folder is larger than the previous largest
                If folderSize > largestSize Then
                    largestSize = folderSize
                    largestFolder = dir
                End If
            Next

            ' Display the largest folder and its size
            Label1.Text = $"Largest Folder: {largestFolder}{Environment.NewLine}Size: {FormatSize(largestSize)}"
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    ' Function to get the size of a directory
    Private Function GetDirectorySize(ByVal dir As String) As Long
        Dim size As Long = 0

        Try
            ' Get all files in the directory
            Dim files As String() = Directory.GetFiles(dir, "*.*", SearchOption.TopDirectoryOnly)

            For Each file As String In files
                size += New FileInfo(file).Length
            Next
        Catch ex As UnauthorizedAccessException
            ' Log or handle the access denied exception
            ' You may want to skip this directory or log it for later review
        Catch ex As Exception
            ' Handle other potential exceptions
        End Try

        Return size
    End Function

    ' Function to format the size for display
    Private Function FormatSize(ByVal size As Long) As String
        Dim sizes As String() = {"B", "KB", "MB", "GB"}
        Dim order As Integer = 0
        While size >= 1024 AndAlso order < sizes.Length - 1
            order += 1
            size /= 1024
        End While
        Return String.Format("{0:0.##} {1}", size, sizes(order))
    End Function
End Class
