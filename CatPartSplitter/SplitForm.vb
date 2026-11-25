Option Strict On

Imports System
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Friend NotInheritable Class SplitForm
    Inherits Form

    Private ReadOnly sourceBox As New TextBox With {.Dock = DockStyle.Fill}
    Private ReadOnly keyBox As New TextBox With {.Dock = DockStyle.Fill, .UseSystemPasswordChar = True}
    Private ReadOnly endpointBox As New TextBox With {.Dock = DockStyle.Fill}
    Private ReadOnly outputBox As New TextBox With {.Dock = DockStyle.Fill}
    Private ReadOnly exportButton As New Button With {.Text = "拆分并导出", .AutoSize = True}
    Private ReadOnly statusLabel As New Label With {.AutoSize = True, .Text = "请选择源 CATPart。"}
    Private exportInProgress As Boolean

    Public Sub New()
        Text = "CATIA V5 R2018 CATPart 拆分工具 V02"
        StartPosition = FormStartPosition.CenterScreen
        MinimumSize = New Size(680, 260)
        Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .Padding = New Padding(14), .ColumnCount = 3, .RowCount = 5}
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.Controls.Add(New Label With {.Text = "源 CATPart", .AutoSize = True}, 0, 0)
        layout.Controls.Add(sourceBox, 1, 0)
        Dim browseButton As New Button With {.Text = "浏览", .AutoSize = True}
        AddHandler browseButton.Click, AddressOf BrowseSource
        layout.Controls.Add(browseButton, 2, 0)
        layout.Controls.Add(New Label With {.Text = "DeepSeek API Key", .AutoSize = True}, 0, 1)
        layout.Controls.Add(keyBox, 1, 1)
        layout.SetColumnSpan(keyBox, 2)
        layout.Controls.Add(New Label With {.Text = "API 地址", .AutoSize = True}, 0, 2)
        layout.Controls.Add(endpointBox, 1, 2)
        layout.SetColumnSpan(endpointBox, 2)
        layout.Controls.Add(New Label With {.Text = "输出目录", .AutoSize = True}, 0, 3)
        layout.Controls.Add(outputBox, 1, 3)
        Dim outputBrowseButton As New Button With {.Text = "浏览", .AutoSize = True}
        AddHandler outputBrowseButton.Click, AddressOf BrowseOutput
        layout.Controls.Add(outputBrowseButton, 2, 3)
        layout.Controls.Add(exportButton, 1, 4)
        layout.Controls.Add(statusLabel, 0, 5)
        layout.SetColumnSpan(statusLabel, 3)
        Controls.Add(layout)
        Dim sample = Path.Combine(AppContext.BaseDirectory, "BE13ASSY.CATPart")
        sourceBox.Text = If(File.Exists(sample), sample, String.Empty)
        keyBox.Text = ReadSetting("api.txt")
        Dim savedEndpoint = ReadSetting("api-url.txt")
        endpointBox.Text = If(String.IsNullOrWhiteSpace(savedEndpoint), "https://api.deepseek.com", savedEndpoint)
        outputBox.Text = If(String.IsNullOrWhiteSpace(sourceBox.Text), String.Empty, Path.GetDirectoryName(sourceBox.Text))
        AddHandler exportButton.Click, AddressOf Export
        AddHandler FormClosing, AddressOf PreventCloseDuringExport
    End Sub

    Private Sub BrowseOutput(sender As Object, eventArgs As EventArgs)
        Using dialog As New FolderBrowserDialog()
            If dialog.ShowDialog(Me) = DialogResult.OK Then outputBox.Text = dialog.SelectedPath
        End Using
    End Sub

    Private Sub BrowseSource(sender As Object, eventArgs As EventArgs)
        Using dialog As New OpenFileDialog With {.Filter = "CATPart (*.CATPart)|*.CATPart", .CheckFileExists = True}
            If dialog.ShowDialog(Me) = DialogResult.OK Then
                sourceBox.Text = dialog.FileName
                If String.IsNullOrWhiteSpace(outputBox.Text) Then outputBox.Text = Path.GetDirectoryName(dialog.FileName)
            End If
        End Using
    End Sub

    Private Sub Export(sender As Object, eventArgs As EventArgs)
        If Not File.Exists(sourceBox.Text) Then
            statusLabel.Text = "请选择存在的 CATPart 文件。"
            Return
        End If
        If String.IsNullOrWhiteSpace(outputBox.Text) Then
            statusLabel.Text = "请输入输出目录。"
            Return
        End If
        Dim sourcePath = Path.GetFullPath(sourceBox.Text)
        Dim outputPath As String
        Try
            outputPath = Path.GetFullPath(outputBox.Text)
        Catch ex As Exception When TypeOf ex Is ArgumentException OrElse TypeOf ex Is IOException OrElse TypeOf ex Is NotSupportedException
            statusLabel.Text = "输出目录无效: " & ex.Message
            Return
        End Try
        Dim apiKey = keyBox.Text.Trim()
        Dim apiUrl = endpointBox.Text.Trim()
        If String.IsNullOrWhiteSpace(apiUrl) Then apiUrl = "https://api.deepseek.com"
        exportInProgress = True
        exportButton.Enabled = False
        statusLabel.Text = "正在连接 CATIA V5 R2018 并导出，请等待。"
        Dim worker As New Threading.Thread(Sub()
                                               Environment.SetEnvironmentVariable("DEEPSEEK_API_KEY", apiKey)
                                               Environment.SetEnvironmentVariable("DEEPSEEK_BASE_URL", apiUrl)
                                               Dim exitCode = Program.Execute(New Program.CommandLineOptions With {.SourcePath = sourcePath, .OutputParentPath = outputPath})
                                               BeginInvoke(Sub()
                                                               Dim settingsError = If(exitCode = 0, SaveSettings(apiKey, apiUrl), String.Empty)
                                                               exportInProgress = False
                                                               exportButton.Enabled = True
                                                               statusLabel.Text = If(exitCode = 0, If(String.IsNullOrEmpty(settingsError), "导出完成。", settingsError), "导出失败: " & Program.LastError)
                                                           End Sub)
                                           End Sub)
        worker.SetApartmentState(Threading.ApartmentState.STA)
        worker.IsBackground = True
        worker.Start()
    End Sub

    Private Shared Function ReadSetting(fileName As String) As String
        Try
            Dim settingPath = IO.Path.Combine(AppContext.BaseDirectory, fileName)
            Return If(File.Exists(settingPath), File.ReadAllText(settingPath, Encoding.UTF8).Trim(), String.Empty)
        Catch
            Return String.Empty
        End Try
    End Function

    Private Function SaveSettings(apiKey As String, apiUrl As String) As String
        Try
            If Not String.IsNullOrWhiteSpace(apiKey) Then File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "api.txt"), apiKey, New UTF8Encoding(False))
            File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "api-url.txt"), apiUrl, New UTF8Encoding(False))
            Return String.Empty
        Catch exception As IOException
            Return "导出完成，但无法保存 API 设置: " & exception.Message
        Catch exception As UnauthorizedAccessException
            Return "导出完成，但无权保存 API 设置: " & exception.Message
        End Try
    End Function

    Private Sub PreventCloseDuringExport(sender As Object, eventArgs As FormClosingEventArgs)
        If exportInProgress Then
            eventArgs.Cancel = True
            statusLabel.Text = "导出正在进行，请等待完成。"
        End If
    End Sub
End Class
