Option Strict Off

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.Json
Imports Microsoft.VisualBasic
Imports System.Windows.Forms

Module Program
    Private Const RequiredRelease As String = "28"
    Private Const NoShowAttribute As Integer = 1
    Private Const ColorInheritanceMode As Integer = 1
    Private Const GlossaryResourceName As String = "CatPartSplitter.Glossary.md"
    Private ReadOnly JsonOptions As New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True}
    Friend Property LastError As String = String.Empty

    <DllImport("oleaut32.dll")>
    Private Function GetActiveObject(ByRef classId As Guid, reserved As IntPtr, <Out, MarshalAs(UnmanagedType.IUnknown)> ByRef activeObject As Object) As Integer
    End Function

    <STAThread>
    Public Function Main(arguments As String()) As Integer
        If arguments.Length = 0 Then
            Application.EnableVisualStyles()
            Application.Run(New SplitForm())
            Return 0
        End If
        Return Execute(CommandLineOptions.Parse(arguments))
    End Function

    Friend Function Execute(options As CommandLineOptions) As Integer
        Try
            LastError = String.Empty
            If options.ShowHelp Then
                PrintUsage()
                Return 0
            End If

            Dim catia As Object = ConnectToRunningCatia()
            VerifyRelease(catia)

            Dim sourceDocument As Object = FindOrOpenPartDocument(catia, options.SourcePath)
            Dim groups = ReadGroups(sourceDocument)
            If groups.Count = 0 Then Throw New InvalidOperationException("源 CATPart 中没有可导出的 PartBody。")

            If options.InspectOnly Then
                PrintGroups(groups)
                Return 0
            End If

            Dim nameMap = GetEnglishNamesAsync(groups.Select(Function(group) group.OriginalName).Distinct(StringComparer.Ordinal).ToList(), options).GetAwaiter().GetResult()
            Dim outputDirectory = CreateOutputDirectory(options.SourcePath, options.OutputParentPath)
            Dim savedParts As New List(Of ExportedPart)()
            Dim manifestItems As New List(Of ManifestItem)()
            Dim colors As New GroupColorGenerator()

            For groupIndex As Integer = 0 To groups.Count - 1
                Dim group = groups(groupIndex)
                Dim filename = nameMap(group.OriginalName)
                Dim partPath = Path.Combine(outputDirectory, filename & ".CATPart")
                Dim color = colors.NextColor(groupIndex)
                Dim createdPart = CreatePart(catia, sourceDocument, group, color, partPath)
                savedParts.Add(New ExportedPart With {.Path = partPath, .OriginalName = group.OriginalName})
                manifestItems.Add(New ManifestItem With {
                    .OriginalName = group.OriginalName,
                    .EnglishFileName = filename & ".CATPart",
                    .SourceBodyCount = group.Bodies.Count,
                    .OutputBodyCount = createdPart.OutputBodyCount
                })
                Console.WriteLine($"已保存 {partPath}")
            Next

            Dim productPath = CreateProduct(catia, outputDirectory, Path.GetFileNameWithoutExtension(options.SourcePath), savedParts)
            File.WriteAllText(Path.Combine(outputDirectory, "split-manifest.json"), JsonSerializer.Serialize(New SplitManifest With {
                .SourcePath = options.SourcePath,
                .ProductFile = Path.GetFileName(productPath),
                .Items = manifestItems
            }, New JsonSerializerOptions With {.WriteIndented = True}), New UTF8Encoding(False))
            Console.WriteLine($"已保存总装 {productPath}")
            Return 0
        Catch exception As Exception
            LastError = exception.Message
            Console.Error.WriteLine("失败: " & exception.Message)
            Return 1
        End Try
    End Function

    Private Function ConnectToRunningCatia() As Object
        Try
            Dim classId As New Guid("87fd6f40-e252-11d5-8040-0010b5fa1031")
            Dim catia As Object = Nothing
            Dim hresult = GetActiveObject(classId, IntPtr.Zero, catia)
            If hresult < 0 Then Marshal.ThrowExceptionForHR(hresult)
            If catia Is Nothing Then Throw New COMException("ROT 中没有 CATIA.Application。")
            Return catia
        Catch exception As Exception
            Throw New InvalidOperationException("未找到已运行的 CATIA V5 R2018 会话。请先启动 R2018，再运行本程序。程序不会创建 CATIA 实例。COM 错误: " & exception.Message, exception)
        End Try
    End Function

    Private Sub VerifyRelease(catia As Object)
        Dim release As String
        Try
            Dim configuration = GetComProperty(catia, "SystemConfiguration")
            release = CStr(GetComProperty(configuration, "Release"))
        Catch exception As Exception
            Throw New InvalidOperationException("无法读取 CATIA.SystemConfiguration.Release，无法确认正在运行的 CATIA 版本。", exception)
        End Try
        If Not String.Equals(release, RequiredRelease, StringComparison.OrdinalIgnoreCase) Then
            Throw New InvalidOperationException("当前 CATIA 版本为 [" & release & "]，需要 CATIA V5 R2018 [" & RequiredRelease & "]。请关闭 R2016 并启动 R2018 后重试。")
        End If
    End Sub

    Private Function GetComProperty(target As Object, propertyName As String) As Object
        Return target.GetType().InvokeMember(propertyName, Reflection.BindingFlags.GetProperty, Nothing, target, Nothing)
    End Function

    Private Function InvokeComMethod(target As Object, methodName As String, ParamArray arguments As Object()) As Object
        Return target.GetType().InvokeMember(methodName, Reflection.BindingFlags.InvokeMethod Or Reflection.BindingFlags.GetProperty, Nothing, target, arguments)
    End Function

    Private Function FindOrOpenPartDocument(catia As Object, sourcePath As String) As Object
        For index As Integer = 1 To CInt(catia.Documents.Count)
            Dim document = catia.Documents.Item(index)
            Try
                If String.Equals(CStr(document.FullName), sourcePath, StringComparison.OrdinalIgnoreCase) Then Return document
            Catch
            End Try
        Next
        Return catia.Documents.Open(sourcePath)
    End Function

    Private Function ReadGroups(sourceDocument As Object) As List(Of BodyGroup)
        Dim part = sourceDocument.Part
        Dim result As New Dictionary(Of String, BodyGroup)(StringComparer.Ordinal)
        For index As Integer = 1 To CInt(part.Bodies.Count)
            Dim body = part.Bodies.Item(index)
            If IsExportable(body) Then
                Dim originalName = CStr(body.Name)
                If originalName.Length = 0 Then originalName = "UnnamedBody"
                Dim group As BodyGroup = Nothing
                If Not result.TryGetValue(originalName, group) Then
                    group = New BodyGroup With {.OriginalName = originalName}
                    result.Add(originalName, group)
                End If
                group.Bodies.Add(body)
            End If
        Next
        Return result.Values.OrderBy(Function(group) group.OriginalName, StringComparer.Ordinal).ToList()
    End Function

    Private Function IsExportable(body As Object) As Boolean
        Try
            If CBool(body.InBooleanOperation) Then Return False
            Return CInt(body.Shapes.Count) > 0
        Catch exception As Exception
            Throw New InvalidOperationException("无法检查实体 [" & CStr(body.Name) & "]。", exception)
        End Try
    End Function

    Private Function CreatePart(catia As Object, sourceDocument As Object, group As BodyGroup, color As RgbColor, partPath As String) As CreatedPart
        Dim document = catia.Documents.Add("Part")
        Dim part = document.Part
        document.Product.PartNumber = group.OriginalName
        Dim outputCount As Integer = 0
        For sourceIndex As Integer = 0 To group.Bodies.Count - 1
            Dim targetBody As Object = Nothing
            Dim copied As Boolean = False
            For copyAttempt As Integer = 1 To 3
                Dim existingBodyCount = CInt(part.Bodies.Count)
                CopyBodyAsResult(sourceDocument, group.Bodies(sourceIndex), document, part)
                part.Update()
                Dim updatedBodyCount = CInt(part.Bodies.Count)
                If updatedBodyCount = existingBodyCount + 1 Then
                    targetBody = part.Bodies.Item(existingBodyCount + 1)
                    copied = True
                    Exit For
                End If
                If updatedBodyCount <> existingBodyCount Then
                    Throw New InvalidOperationException("粘贴实体时 Body 数量异常。分组 [" & group.OriginalName & "]，实体 " & (sourceIndex + 1).ToString(CultureInfo.InvariantCulture) & "，粘贴前 " & existingBodyCount.ToString(CultureInfo.InvariantCulture) & "，粘贴后 " & updatedBodyCount.ToString(CultureInfo.InvariantCulture) & "。")
                End If
                If copyAttempt < 3 Then Threading.Thread.Sleep(250)
            Next
            If Not copied Then
                Throw New InvalidOperationException("粘贴实体后 Body 数量未变化。分组 [" & group.OriginalName & "]，实体 " & (sourceIndex + 1).ToString(CultureInfo.InvariantCulture) & "，已重试 3 次。")
            End If
            targetBody.Name = group.OriginalName
            If CInt(targetBody.Shapes.Count) < 1 Then
                Throw New InvalidOperationException("复制 [" & group.OriginalName & "] 的第 " & (sourceIndex + 1).ToString(CultureInfo.InvariantCulture) & " 个实体后没有生成形状。输出已保留在 " & partPath & " 的同级目录。")
            End If
            ApplyColor(document, targetBody, color)
            outputCount += 1
        Next
        HideOriginPlanes(document, part)
        document.SaveAs(partPath)
        Return New CreatedPart With {.OutputBodyCount = outputCount}
    End Function

    Private Sub ApplyColor(document As Object, body As Object, color As RgbColor)
        Try
            InvokeComMethod(document, "Activate")
            Dim selection = GetComProperty(document, "Selection")
            InvokeComMethod(selection, "Clear")
            Try
                Dim shapes = GetComProperty(body, "Shapes")
                For shapeIndex As Integer = 1 To CInt(GetComProperty(shapes, "Count"))
                    InvokeComMethod(selection, "Add", InvokeComMethod(shapes, "Item", shapeIndex))
                Next
                Dim visualProperties = GetComProperty(selection, "VisProperties")
                InvokeComMethod(visualProperties, "SetRealColor", color.Red, color.Green, color.Blue, ColorInheritanceMode)
            Finally
                InvokeComMethod(selection, "Clear")
            End Try
        Catch exception As Exception
            Throw New InvalidOperationException("无法为实体 [" & CStr(body.Name) & "] 设置随机颜色。", exception)
        End Try
    End Sub

    Private Sub HideOriginPlanes(document As Object, part As Object)
        Try
            InvokeComMethod(document, "Activate")
            Dim selection = GetComProperty(document, "Selection")
            InvokeComMethod(selection, "Clear")
            Try
                Dim originElements = GetComProperty(part, "OriginElements")
                InvokeComMethod(selection, "Add", GetComProperty(originElements, "PlaneXY"))
                InvokeComMethod(selection, "Add", GetComProperty(originElements, "PlaneYZ"))
                InvokeComMethod(selection, "Add", GetComProperty(originElements, "PlaneZX"))
                Dim visualProperties = GetComProperty(selection, "VisProperties")
                InvokeComMethod(visualProperties, "SetShow", NoShowAttribute)
            Finally
                InvokeComMethod(selection, "Clear")
            End Try
        Catch exception As Exception
            Throw New InvalidOperationException("无法隐藏导出零件的 XY、YZ、ZX 基准面。", exception)
        End Try
    End Sub

    Private Sub CopyBodyAsResult(sourceDocument As Object, sourceBody As Object, targetDocument As Object, targetPart As Object)
        sourceDocument.Activate()
        Dim sourceSelection = sourceDocument.Selection
        sourceSelection.Clear()
        sourceSelection.Add(sourceBody)
        sourceSelection.Copy()

        targetDocument.Activate()
        Dim targetSelection = targetDocument.Selection
        targetSelection.Clear()
        targetSelection.Add(targetPart)
        targetSelection.PasteSpecial("CATPrtResultWithOutLink")
        targetSelection.Clear()
        sourceSelection.Clear()
    End Sub

    Private Function CreateProduct(catia As Object, outputDirectory As String, sourceStem As String, savedParts As List(Of ExportedPart)) As String
        Dim productDocument = catia.Documents.Add("Product")
        Dim product = productDocument.Product
        Dim productName = SanitizeFileStem(sourceStem) & "_Split"
        product.PartNumber = productName
        For Each savedPart In savedParts
            Dim existingComponentCount = CInt(product.Products.Count)
            product.Products.AddComponentsFromFiles(New Object() {savedPart.Path}, "All")
            Dim instance = product.Products.Item(existingComponentCount + 1)
            instance.Name = savedPart.OriginalName & ".1"
        Next
        product.Update()
        Dim productPath = Path.Combine(outputDirectory, productName & ".CATProduct")
        productDocument.SaveAs(productPath)
        Return productPath
    End Function

    Private Async Function GetEnglishNamesAsync(originalNames As List(Of String), options As CommandLineOptions) As Task(Of Dictionary(Of String, String))
        Dim suppliedMap = ReadProvidedNameMap(options.NameMapPath)
        Dim glossary = LoadGlossary()
        Dim unresolved As New List(Of String)()
        Dim candidates As New Dictionary(Of String, List(Of String))(StringComparer.Ordinal)

        For Each originalName In originalNames
            If suppliedMap.ContainsKey(originalName) Then Continue For

            Dim entries As List(Of GlossaryEntry) = Nothing
            If glossary.TryGetValue(originalName, entries) Then
                Dim exactCandidates = entries.Select(Function(entry) entry.English).Where(Function(english) Not String.IsNullOrWhiteSpace(english)).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                If exactCandidates.Count = 1 Then
                    suppliedMap.Add(originalName, exactCandidates(0))
                    Continue For
                End If
                If exactCandidates.Count > 1 Then candidates.Add(originalName, exactCandidates)
            End If
            unresolved.Add(originalName)
        Next

        If unresolved.Count > 0 Then
            Dim apiKey = Environment.GetEnvironmentVariable("DEEPSEEK_API_KEY")
            If String.IsNullOrWhiteSpace(apiKey) Then
                Throw New InvalidOperationException("术语库没有唯一匹配以下名称，且未设置 DEEPSEEK_API_KEY: " & String.Join("、", unresolved) & "。请设置 Key 或使用 --name-map 指定英文名称。")
            End If
            Dim translated = Await TranslateWithDeepSeekAsync(unresolved, apiKey, Environment.GetEnvironmentVariable("DEEPSEEK_BASE_URL"), candidates, glossary, FindRelevantGlossaryEntries(unresolved, glossary))
            For Each item In translated
                suppliedMap(item.Key) = item.Value
            Next
        End If
        Return MakeUniqueFileStems(originalNames, suppliedMap)
    End Function

    Private Function ReadProvidedNameMap(path As String) As Dictionary(Of String, String)
        If String.IsNullOrWhiteSpace(path) Then Return New Dictionary(Of String, String)(StringComparer.Ordinal)
        If Not File.Exists(path) Then Throw New FileNotFoundException("找不到 --name-map 文件。", path)
        Dim map = JsonSerializer.Deserialize(Of Dictionary(Of String, String))(File.ReadAllText(path), JsonOptions)
        If map Is Nothing Then Throw New InvalidOperationException("--name-map 必须是 JSON 对象。")
        Return New Dictionary(Of String, String)(map, StringComparer.Ordinal)
    End Function

    Private Async Function TranslateWithDeepSeekAsync(names As List(Of String), apiKey As String, suppliedBaseUrl As String, exactCandidates As Dictionary(Of String, List(Of String)), glossary As Dictionary(Of String, List(Of GlossaryEntry)), relevantEntries As List(Of GlossaryEntry)) As Task(Of Dictionary(Of String, String))
        Dim baseUrl = If(String.IsNullOrWhiteSpace(suppliedBaseUrl), "https://api.deepseek.com", suppliedBaseUrl.TrimEnd("/"c))
        Dim endpoint As Uri = Nothing
        If Not Uri.TryCreate(baseUrl, UriKind.Absolute, endpoint) OrElse endpoint.Scheme <> Uri.UriSchemeHttps OrElse Not String.IsNullOrEmpty(endpoint.UserInfo) Then
            Throw New InvalidOperationException("DeepSeek API 地址必须是没有用户名或密码的绝对 HTTPS 地址。")
        End If
        Dim prompt = "Translate every CAD body name into a concise English HVAC engineering name. Return JSON only in this exact form: {""translations"":[{""original"":""..."",""english"":""...""}]}. Preserve every original exactly and include each input once. For a name in Exact glossary candidates, choose one listed English value verbatim. For other names, prefer the supplied relevant HVAC glossary terminology where applicable. English values must use ASCII letters, digits, spaces, hyphens, underscores, or periods only. Input names: " & JsonSerializer.Serialize(names) & ". Exact glossary candidates: " & FormatExactCandidates(exactCandidates, glossary) & ". Relevant HVAC glossary entries: " & JsonSerializer.Serialize(relevantEntries.Select(Function(entry) New With {.chinese = entry.Chinese, .english = entry.English}).ToList())
        Dim requestBody = New With {
            .model = "deepseek-flash",
            .messages = New Object() {New With {.role = "system", .content = "Return JSON only."}, New With {.role = "user", .content = prompt}},
            .response_format = New With {.type = "json_object"},
            .temperature = 0,
            .max_tokens = 2048,
            .thinking = New With {.type = "disabled"}
        }
        Using handler As New HttpClientHandler With {.AllowAutoRedirect = False}
        Using client As New HttpClient(handler)
            client.DefaultRequestHeaders.Authorization = New Headers.AuthenticationHeaderValue("Bearer", apiKey)
            Using response = Await client.PostAsync(baseUrl & "/chat/completions", New StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json"))
                Dim raw = Await response.Content.ReadAsStringAsync()
                If Not response.IsSuccessStatusCode Then Throw New InvalidOperationException("DeepSeek API 请求失败，HTTP 状态码: " & CInt(response.StatusCode).ToString(CultureInfo.InvariantCulture))
                Dim envelope = JsonDocument.Parse(raw)
                Dim content = envelope.RootElement.GetProperty("choices")(0).GetProperty("message").GetProperty("content").GetString()
                If String.IsNullOrWhiteSpace(content) Then Throw New InvalidOperationException("DeepSeek API 返回了空翻译结果。")
                Dim translation = JsonSerializer.Deserialize(Of TranslationResponse)(content, JsonOptions)
                If translation?.Translations Is Nothing Then Throw New InvalidOperationException("DeepSeek API 翻译结果缺少 translations。")
                Dim result = translation.Translations.ToDictionary(Function(item) item.Original, Function(item) item.English, StringComparer.Ordinal)
                If result.Count <> names.Count OrElse names.Any(Function(name) Not result.ContainsKey(name)) Then Throw New InvalidOperationException("DeepSeek API 未返回每个实体名的一一对应翻译。")
                For Each entry In exactCandidates
                    Dim selected = entry.Value.FirstOrDefault(Function(candidate) String.Equals(candidate, result(entry.Key)?.Trim(), StringComparison.OrdinalIgnoreCase))
                    If selected Is Nothing Then
                        Throw New InvalidOperationException("DeepSeek 未从术语库候选中选出 [" & entry.Key & "] 的英文名称。可重试或用 --name-map 指定。")
                    End If
                    result(entry.Key) = selected
                Next
                Return result
            End Using
        End Using
        End Using
    End Function

    Private Function LoadGlossary() As Dictionary(Of String, List(Of GlossaryEntry))
        Dim result As New Dictionary(Of String, List(Of GlossaryEntry))(StringComparer.Ordinal)
        Dim assembly = Reflection.Assembly.GetExecutingAssembly()
        Using stream = assembly.GetManifestResourceStream(GlossaryResourceName)
            If stream Is Nothing Then Throw New InvalidOperationException("程序内未找到空调零部件中英术语库资源。")
            Using reader As New StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                Dim category As String = String.Empty
                While Not reader.EndOfStream
                    Dim line = reader.ReadLine()
                    If line.StartsWith("## ", StringComparison.Ordinal) Then
                        category = line.Substring(3).Trim()
                        Continue While
                    End If
                    If Not line.StartsWith("|", StringComparison.Ordinal) OrElse line.Contains("---", StringComparison.Ordinal) Then Continue While
                    Dim columns = line.Split("|"c).Select(Function(column) column.Trim()).ToArray()
                    If columns.Length < 5 OrElse String.Equals(columns(1), "编号", StringComparison.Ordinal) Then Continue While
                    Dim english = columns(2)
                    Dim chinese = columns(4)
                    If String.IsNullOrWhiteSpace(english) OrElse String.IsNullOrWhiteSpace(chinese) Then Continue While
                    Dim entries As List(Of GlossaryEntry) = Nothing
                    If Not result.TryGetValue(chinese, entries) Then
                        entries = New List(Of GlossaryEntry)()
                        result.Add(chinese, entries)
                    End If
                    entries.Add(New GlossaryEntry With {.Chinese = chinese, .English = english, .Category = category, .Material = If(columns.Length > 5, columns(5), String.Empty), .PartNumber = If(columns.Length > 6, columns(6), String.Empty)})
                End While
            End Using
        End Using
        Return result
    End Function

    Private Function FormatExactCandidates(exactCandidates As Dictionary(Of String, List(Of String)), glossary As Dictionary(Of String, List(Of GlossaryEntry))) As String
        Dim result As New List(Of Object)()
        For Each pair In exactCandidates
            Dim glossaryEntries As List(Of GlossaryEntry) = Nothing
            glossary.TryGetValue(pair.Key, glossaryEntries)
            result.Add(New With {
                .original = pair.Key,
                .candidates = glossaryEntries.Where(Function(entry) pair.Value.Contains(entry.English, StringComparer.OrdinalIgnoreCase)).Select(Function(entry) New With {.english = entry.English, .category = entry.Category, .material = entry.Material, .partNumber = entry.PartNumber}).ToList()
            })
        Next
        Return JsonSerializer.Serialize(result)
    End Function

    Private Function FindRelevantGlossaryEntries(names As List(Of String), glossary As Dictionary(Of String, List(Of GlossaryEntry))) As List(Of GlossaryEntry)
        Dim result As New List(Of GlossaryEntry)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim matchesByName As New List(Of List(Of GlossaryEntry))()
        For Each name In names
            Dim matches = glossary.Values.SelectMany(Function(entries) entries).
                Select(Function(entry) New With {.Entry = entry, .Score = LongestSharedGlossaryTerm(name, entry.Chinese)}).
                Where(Function(match) match.Score >= 2).
                OrderByDescending(Function(match) match.Score).
                ThenBy(Function(match) match.Entry.Chinese, StringComparer.Ordinal).
                ThenBy(Function(match) match.Entry.English, StringComparer.Ordinal).
                Take(4).
                Select(Function(match) match.Entry).
                ToList()
            matchesByName.Add(matches)
        Next

        For matchIndex As Integer = 0 To 3
            For Each matches In matchesByName
                If matchIndex >= matches.Count Then Continue For
                Dim entry = matches(matchIndex)
                Dim key = entry.Chinese & ChrW(0) & entry.English
                If seen.Add(key) Then result.Add(entry)
                If result.Count = 32 Then Return result
            Next
        Next
        Return result
    End Function

    Private Function LongestSharedGlossaryTerm(name As String, glossaryName As String) As Integer
        Dim longest As Integer = 0
        For length As Integer = Math.Min(4, name.Length) To 2 Step -1
            For startIndex As Integer = 0 To name.Length - length
                If glossaryName.Contains(name.Substring(startIndex, length), StringComparison.Ordinal) Then
                    longest = length
                    Exit For
                End If
            Next
            If longest > 0 Then Exit For
        Next
        Return longest
    End Function

    Private Function MakeUniqueFileStems(originalNames As List(Of String), translations As Dictionary(Of String, String)) As Dictionary(Of String, String)
        Dim used As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim result As New Dictionary(Of String, String)(StringComparer.Ordinal)
        For Each originalName In originalNames.OrderBy(Function(name) name, StringComparer.Ordinal)
            If Not translations.ContainsKey(originalName) Then Throw New InvalidOperationException("缺少 [" & originalName & "] 的英文名称。")
            Dim baseStem = SanitizeFileStem(translations(originalName))
            If IsReservedWindowsName(baseStem) Then baseStem = "_" & baseStem
            Dim stem = baseStem
            Dim suffix As Integer = 2
            While used.Contains(stem)
                stem = baseStem & "_" & suffix.ToString(CultureInfo.InvariantCulture)
                suffix += 1
            End While
            used.Add(stem)
            result.Add(originalName, stem)
        Next
        Return result
    End Function

    Private Function SanitizeFileStem(value As String) As String
        value = value.Replace("×", "x")
        Dim text = New String(value.Where(Function(character) character <= ChrW(127) AndAlso (Char.IsLetterOrDigit(character) OrElse character = " "c OrElse character = "-"c OrElse character = "_"c OrElse character = "."c)).ToArray()).Trim(" "c, "."c)
        text = String.Join("_", text.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries))
        If text.Length = 0 Then Throw New InvalidOperationException("翻译没有产生可用的英文 ASCII 文件名。")
        Return text
    End Function

    Private Function IsReservedWindowsName(stem As String) As Boolean
        Dim upper = stem.Split("."c)(0).ToUpperInvariant()
        If upper = "CON" OrElse upper = "PRN" OrElse upper = "AUX" OrElse upper = "NUL" Then Return True
        If upper.StartsWith("COM", StringComparison.Ordinal) OrElse upper.StartsWith("LPT", StringComparison.Ordinal) Then
            Dim suffix = upper.Substring(3)
            Dim portNumber As Integer
            Return Integer.TryParse(suffix, portNumber) AndAlso portNumber >= 1 AndAlso portNumber <= 9
        End If
        Return False
    End Function

    Private Function CreateOutputDirectory(sourcePath As String, outputParentPath As String) As String
        Dim parent = If(String.IsNullOrWhiteSpace(outputParentPath), Path.GetDirectoryName(sourcePath), outputParentPath)
        Dim sourceName = SanitizeFileStem(Path.GetFileNameWithoutExtension(sourcePath))
        For serial As Integer = 1 To 999
            Dim candidate = Path.Combine(parent, $"{sourceName}_Split_{DateTime.Now:yyyyMMdd_HHmmss}_{serial:000}")
            If Not Directory.Exists(candidate) Then
                Directory.CreateDirectory(candidate)
                Return candidate
            End If
        Next
        Throw New IOException("无法创建一个新的输出目录。")
    End Function

    Private Sub PrintGroups(groups As List(Of BodyGroup))
        Console.WriteLine($"可导出组数: {groups.Count}")
        For Each group In groups
            Console.WriteLine($"{group.OriginalName}    实体数: {group.Bodies.Count}")
        Next
    End Sub

    Private Sub PrintUsage()
        Console.WriteLine("CatPartSplitter [source.CATPart] [--inspect] [--name-map path.json] [--output directory]")
        Console.WriteLine("DEEPSEEK_API_KEY is required unless every name is supplied by --name-map.")
    End Sub

    Friend NotInheritable Class CommandLineOptions
        Public Property SourcePath As String = Path.GetFullPath("BE13ASSY.CATPart")
        Public Property InspectOnly As Boolean
        Public Property NameMapPath As String = String.Empty
        Public Property OutputParentPath As String = String.Empty
        Public Property ShowHelp As Boolean

        Public Shared Function Parse(arguments As String()) As CommandLineOptions
            Dim options As New CommandLineOptions()
            Dim sourceSet As Boolean
            For index As Integer = 0 To arguments.Length - 1
                Select Case arguments(index)
                    Case "--inspect" : options.InspectOnly = True
                    Case "--name-map"
                        index += 1
                        If index >= arguments.Length Then Throw New ArgumentException("--name-map 缺少文件路径。")
                        options.NameMapPath = Path.GetFullPath(arguments(index))
                    Case "--output"
                        index += 1
                        If index >= arguments.Length Then Throw New ArgumentException("--output 缺少目录路径。")
                        options.OutputParentPath = Path.GetFullPath(arguments(index))
                    Case "--help", "-h", "/?" : options.ShowHelp = True
                    Case Else
                        If sourceSet OrElse arguments(index).StartsWith("-", StringComparison.Ordinal) Then Throw New ArgumentException("未知参数: " & arguments(index))
                        options.SourcePath = Path.GetFullPath(arguments(index))
                        sourceSet = True
                End Select
            Next
            If Not options.ShowHelp AndAlso Not File.Exists(options.SourcePath) Then Throw New FileNotFoundException("找不到源 CATPart。", options.SourcePath)
            Return options
        End Function
    End Class

    Private NotInheritable Class BodyGroup
        Public Property OriginalName As String = String.Empty
        Public ReadOnly Property Bodies As New List(Of Object)()
    End Class

    Private NotInheritable Class CreatedPart
        Public Property OutputBodyCount As Integer
    End Class

    Private NotInheritable Class ExportedPart
        Public Property Path As String = String.Empty
        Public Property OriginalName As String = String.Empty
    End Class

    Private NotInheritable Class RgbColor
        Public Property Red As Integer
        Public Property Green As Integer
        Public Property Blue As Integer
    End Class

    Private NotInheritable Class GroupColorGenerator
        Private ReadOnly random As New Random()
        Private ReadOnly startingHue As Double

        Public Sub New()
            startingHue = random.NextDouble() * 360.0
        End Sub

        Public Function NextColor(groupIndex As Integer) As RgbColor
            Dim hue = (startingHue + groupIndex * 137.508) Mod 360.0
            If hue < 25.0 Then hue += 45.0
            If hue > 335.0 Then hue -= 40.0
            Dim saturation = 0.6 + random.NextDouble() * 0.15
            Dim value = 0.88 + random.NextDouble() * 0.08
            Return HsvToRgb(hue, saturation, value)
        End Function

        Private Shared Function HsvToRgb(hue As Double, saturation As Double, value As Double) As RgbColor
            Dim sector = CInt(Math.Floor(hue / 60.0)) Mod 6
            Dim fraction = hue / 60.0 - Math.Floor(hue / 60.0)
            Dim p = value * (1.0 - saturation)
            Dim q = value * (1.0 - fraction * saturation)
            Dim t = value * (1.0 - (1.0 - fraction) * saturation)
            Dim red As Double
            Dim green As Double
            Dim blue As Double
            Select Case sector
                Case 0 : red = value : green = t : blue = p
                Case 1 : red = q : green = value : blue = p
                Case 2 : red = p : green = value : blue = t
                Case 3 : red = p : green = q : blue = value
                Case 4 : red = t : green = p : blue = value
                Case Else : red = value : green = p : blue = q
            End Select
            Return New RgbColor With {
                .Red = CInt(Math.Floor(red * 255.0)),
                .Green = CInt(Math.Floor(green * 255.0)),
                .Blue = CInt(Math.Floor(blue * 255.0))
            }
        End Function
    End Class

    Private NotInheritable Class TranslationResponse
        Public Property Translations As List(Of TranslationItem)
    End Class

    Private NotInheritable Class GlossaryEntry
        Public Property Chinese As String = String.Empty
        Public Property English As String = String.Empty
        Public Property Category As String = String.Empty
        Public Property Material As String = String.Empty
        Public Property PartNumber As String = String.Empty
    End Class

    Private NotInheritable Class TranslationItem
        Public Property Original As String = String.Empty
        Public Property English As String = String.Empty
    End Class

    Private NotInheritable Class SplitManifest
        Public Property SourcePath As String = String.Empty
        Public Property ProductFile As String = String.Empty
        Public Property Items As List(Of ManifestItem) = New List(Of ManifestItem)()
    End Class

    Private NotInheritable Class ManifestItem
        Public Property OriginalName As String = String.Empty
        Public Property EnglishFileName As String = String.Empty
        Public Property SourceBodyCount As Integer
        Public Property OutputBodyCount As Integer
    End Class
End Module
