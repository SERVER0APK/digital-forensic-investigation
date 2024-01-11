Imports System.IO
Imports System.Management
Imports System.Net.NetworkInformation
Imports Microsoft.Win32
Imports Microsoft.Win32.TaskScheduler


Public Class info
    Private Sub info_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Get System Information
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        Dim collection As ManagementObjectCollection = searcher.Get()

        'Print system information
        Try
            'Specify device information
            For Each info As ManagementObject In collection
                'Join the manufacturer's name and device model
                TextBox4.Text = $"{info("Manufacturer")} {info("Model")}"
            Next
        Catch ex As Exception
            MessageBox.Show($"An error occurred while loading device information: {ex.Message}", "Wrong", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ControlExtension.Draggable(Me, True)

        Dim installDate As String = GetOriginalInstallDate()
        TextBox1.Text = installDate
        Dim osInfo As String = GetOperatingSystemInfo()
        TextBox2.Text = osInfo
        TextBox3.Text = Environment.MachineName
        Dim psi As New ProcessStartInfo()
        psi.FileName = "powershell.exe"
        psi.Arguments = "Get-MpThreatDetection"
        psi.RedirectStandardOutput = True
        psi.UseShellExecute = False
        psi.CreateNoWindow = True
        Dim process As New Process()
        process.StartInfo = psi
        process.Start()
        Dim output As String = process.StandardOutput.ReadToEnd()
        process.WaitForExit()
        RichTextBox2.Text = output
        USBS()
        STARTUP()
        DisplayScheduledTasks()
        wifi()
    End Sub
    Sub wifi()
        'Call the get grids function and display it in RichTextBox5
        DisplayWiFiNetworks(RichTextBox5)
    End Sub

    Sub DisplayWiFiNetworks(textBox As RichTextBox)
        Dim networks As List(Of String) = GetWiFiNetworks()

        'Extract passwords and display them in RichTextBox5
        For Each network In networks
            Dim password As String = GetWiFiPassword(network)
            textBox.AppendText($"Network: {network}, Password: {password}{Environment.NewLine}")
        Next
    End Sub

    Function GetWiFiNetworks() As List(Of String)
        Dim networks As New List(Of String)()

        'First command: display networks
        Dim processStartInfo As New ProcessStartInfo() With {
        .FileName = "netsh",
        .Arguments = "wlan show profiles",
        .RedirectStandardOutput = True,
        .UseShellExecute = False,
        .CreateNoWindow = True
    }

        Using process As Process = Process.Start(processStartInfo)
            Using reader As StreamReader = process.StandardOutput
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    If line.Contains("All User Profile") Then
                        networks.Add(line.Split(":")(1).Trim())
                    End If
                End While
            End Using
        End Using

        Return networks
    End Function

    Function GetWiFiPassword(networkName As String) As String
        'Second command: password extraction
        Dim processStartInfo As New ProcessStartInfo() With {
        .FileName = "netsh",
        .Arguments = $"wlan show profile name=""{networkName}"" key=clear",
        .RedirectStandardOutput = True,
        .UseShellExecute = False,
        .CreateNoWindow = True
    }

        Using process As Process = Process.Start(processStartInfo)
            Using reader As StreamReader = process.StandardOutput
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    If line.Contains("Key Content") Then
                        Return line.Split(":")(1).Trim()
                    End If
                End While
            End Using
        End Using

        Return "Not available"
    End Function

    Sub STARTUP()
        'Specify the registry path for startup programs
        Dim keyPath As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

        'Open the registry key
        Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(keyPath)

        'Checking if the key exists
        If key IsNot Nothing Then
            'Get the names of the values (programs)
            Dim valueNames As String() = key.GetValueNames()

            'View information about each program
            For Each valueName In valueNames
                ' اسم الملف
                Dim fileName As String = key.GetValue(valueName).ToString()

                'Program installation date
                Dim installDate As Date = GetInstallDate(fileName)

                'Add information to RichTextBox instead of resetting it
                RichTextBox3.AppendText("File name: " & fileName & vbCrLf)
                RichTextBox3.AppendText("File path: " & GetFilePath(fileName) & vbCrLf)
                RichTextBox3.AppendText("Date of installation: " & installDate.ToString() & vbCrLf)
                RichTextBox3.AppendText("------------------------" & vbCrLf)
            Next
        Else
            'If the key is not found
            RichTextBox3.Text = "The registry key could not be found."
        End If

    End Sub


    'The function for obtaining the path of the program file
    Function GetFilePath(fileName As String) As String
        Try
            Dim fileInfo As New System.IO.FileInfo(fileName)
            Return fileInfo.FullName
        Catch ex As Exception
            Return "The file cannot be found"
        End Try
    End Function

    'The function for obtaining the installation date of the program
    Function GetInstallDate(fileName As String) As Date
        Try
            Dim fileInfo As New System.IO.FileInfo(fileName)
            Return fileInfo.CreationTime
        Catch ex As Exception
            Return Date.MinValue
        End Try
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Create a directory to store section files
        Dim sectionDir As String = Path.Combine(Application.StartupPath, "Sections")
        If Not Directory.Exists(sectionDir) Then
            Directory.CreateDirectory(sectionDir)
        End If

        ' Create separate HTML files for each section
        CreateSectionFile("Installation Date of the Operating System", TextBox1.Text, Path.Combine(sectionDir, "InstallationDate.html"))
        CreateSectionFile("System Type and Kernel", TextBox2.Text, Path.Combine(sectionDir, "SystemType.html"))
        CreateSectionFile("Computer Name", TextBox3.Text, Path.Combine(sectionDir, "ComputerName.html"))
        CreateSectionFile("Manufacturer Name", TextBox4.Text, Path.Combine(sectionDir, "ManufacturerName.html"))
        CreateSectionFile("WiFi Networks", RichTextBox5.Text, Path.Combine(sectionDir, "WiFiNetworks.html"), True)
        CreateSectionFile("USB Names", RichTextBox1.Text, Path.Combine(sectionDir, "USBNames.html"), True)
        CreateSectionFile("Detected Viruses", RichTextBox2.Text, Path.Combine(sectionDir, "DetectedViruses.html"))
        CreateSectionFile("StartUP", RichTextBox3.Text, Path.Combine(sectionDir, "StartUP.html"), True)
        CreateSectionFile("Tasks", RichTextBox4.Text, Path.Combine(sectionDir, "Tasks.html"), True)

        ' Create the main HTML file with buttons for each section
        CreateMainPage(sectionDir)

        MessageBox.Show("Information saved successfully in the HTML files.")
    End Sub

    Private Sub CreateSectionFile(sectionTitle As String, content As String, filePath As String, Optional usePre As Boolean = False, Optional useDoubleLineBreaks As Boolean = True)
        ' Open the section file for writing
        Using writer As New StreamWriter(filePath, False)
            ' Write the HTML header
            writer.WriteLine("<html>")
            writer.WriteLine("<head><title>" & sectionTitle & "</title></head>")
            writer.WriteLine("<body>")

            ' Save information to the section file
            WriteHtmlSection(writer, sectionTitle, content, usePre, useDoubleLineBreaks)

            ' Write the HTML footer
            writer.WriteLine("</body>")
            writer.WriteLine("</html>")
        End Using
    End Sub
    Private Sub WriteHtmlSection(writer As StreamWriter, sectionTitle As String, content As String, Optional usePre As Boolean = False, Optional useDoubleLineBreaks As Boolean = True)
        ' Write the section title
        writer.WriteLine($"<h2>{sectionTitle}</h2>")

        ' Split the content into groups (each group represents a result)
        Dim resultGroups() As String = content.Split(New String() {"ActionSuccess"}, StringSplitOptions.RemoveEmptyEntries)

        ' Use <pre> tag if specified
        If usePre Then
            writer.WriteLine("<pre>")
            For Each group As String In resultGroups
                writer.WriteLine($"ActionSuccess{group.Replace(vbCrLf, "<br>").Replace("<br>", "<br>")}")
            Next
            writer.WriteLine("</pre>")
        Else
            ' Write each result group as a separate div
            Dim isFirstGroup As Boolean = True
            For Each group As String In resultGroups
                If Not isFirstGroup Then
                    ' Add a horizontal line as a separator between result groups
                    writer.WriteLine("<hr>")
                Else
                    isFirstGroup = False
                End If

                ' Extract network name and password
                Dim lines() As String = group.Split(Environment.NewLine)
                For Each line As String In lines
                    If Not String.IsNullOrWhiteSpace(line) Then
                        Dim parts() As String = line.Split(","c)
                        If parts.Length = 2 Then
                            Dim networkName As String = parts(0).Trim()
                            Dim password As String = parts(1).Trim()
                            ' Write network name and password in the desired format
                            writer.WriteLine($"<div>Network: {networkName}, Password: {password}</div>")
                        Else
                            ' Write the line as a separate div for other cases
                            writer.WriteLine($"<div>{line.Replace(vbCrLf, "<br>").Replace("<br>", "<br>")}</div>")
                        End If
                    End If
                Next
            Next
        End If

        ' Add a line break for better organization
        If useDoubleLineBreaks Then
            writer.WriteLine("<br><br>")
        Else
            writer.WriteLine("<br>")
        End If
    End Sub

    Private Sub CreateMainPage(sectionDir As String)
        ' Create the main HTML file
        Dim mainFilePath As String = Path.Combine(Application.StartupPath, "MainPage.html")

        ' Open the main HTML file for writing
        Using writer As New StreamWriter(mainFilePath, False)
            ' Write the HTML header
            writer.WriteLine("<html>")
            writer.WriteLine("<head><title>Main Page</title></head>")
            writer.WriteLine("<body>")

            ' Create buttons for each section
            For Each sectionFile As String In Directory.GetFiles(sectionDir)
                Dim sectionName As String = Path.GetFileNameWithoutExtension(sectionFile)
                Dim relativePath As String = GetRelativePath(mainFilePath, sectionFile)
                writer.WriteLine($"<button onclick=""window.open('{relativePath}', '_blank')"">{sectionName}</button>")
            Next

            ' Write the HTML footer
            writer.WriteLine("</body>")
            writer.WriteLine("</html>")
        End Using
    End Sub

    Private Function GetRelativePath(basePath As String, targetPath As String) As String
        Dim baseUri As New Uri(basePath)
        Dim targetUri As New Uri(targetPath)
        Dim relativeUri As Uri = baseUri.MakeRelativeUri(targetUri)
        Return Uri.UnescapeDataString(relativeUri.ToString())
    End Function



    Sub DisplayScheduledTasks()
        'Create a task scheduler
        Using ts As TaskService = New TaskService()
            'Get all the scheduled tasks
            Dim tasks As IEnumerable(Of Task) = ts.AllTasks

            'View information about each scheduled task
            For Each task As Task In tasks
                RichTextBox4.AppendText("Task Name: " & task.Name & vbCrLf)
                RichTextBox4.AppendText("Path: " & task.Path & vbCrLf)
                RichTextBox4.AppendText("State: " & task.State.ToString() & vbCrLf)
                RichTextBox4.AppendText("Next Run Time: " & task.NextRunTime.ToString() & vbCrLf)
                RichTextBox4.AppendText("------------------------" & vbCrLf)
            Next
        End Using
    End Sub
    Sub USBS()
        Dim registryKey2 As RegistryKey = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Enum\USBSTOR", True)
        If registryKey2 Is Nothing Then
            Exit Sub
        End If
        Dim subKeyNames2 As String() = registryKey2.GetSubKeyNames()
        Dim deviceDesc As String
        Dim mfg As String
        Dim i As Integer = 0
        For Each subKeyName2 As String In subKeyNames2
            Dim subKey1 As RegistryKey = registryKey2.OpenSubKey(subKeyName2)
            Dim subKey2Names As String() = subKey1.GetSubKeyNames()
            For Each subKey2Name As String In subKey2Names
                Dim subKey2 As RegistryKey = subKey1.OpenSubKey(subKey2Name)
                deviceDesc = subKey2.GetValue("FriendlyName", "").ToString()
                mfg = subKey2.GetValue("Mfg", "").ToString()
                ' Append the result to RichTextBox1
                RichTextBox1.AppendText(String.Format("FriendlyName: {0}{1}", deviceDesc, Environment.NewLine))
                'RichTextBox1.AppendText(String.Format("Mfg: {0}{1}", mfg, Environment.NewLine))
                RichTextBox1.AppendText("=============" & Environment.NewLine)
                i = i + 1
                subKey2.Close()
            Next
            subKey1.Close()
        Next
        Dim lines() As String = RichTextBox1.Lines
        Dim count As Integer = 0
        For Each line As String In lines
            If line.StartsWith("FriendlyName:") Then
                count += 1
            End If
        Next

    End Sub
    Private Function GetOperatingSystemInfo() As String
        Dim osInfo As String = ""

        Try
            Dim osSearcher As New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
            Dim osMgmtObj As ManagementObject = osSearcher.Get().OfType(Of ManagementObject)().FirstOrDefault()

            If osMgmtObj IsNot Nothing Then
                Dim osName As String = osMgmtObj("Caption").ToString()
                Dim osArchitecture As String = osMgmtObj("OSArchitecture").ToString()

                osInfo = $"{osName}{Environment.NewLine}  Kernel: {osArchitecture}"
                TextBox2.Text = osInfo
            End If

        Catch ex As Exception
            osInfo = "Unknown"
        End Try

        Return osInfo
    End Function
    Private Function GetOriginalInstallDate() As String
        Dim oSearcher As New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
        Dim oMgmtObj As ManagementObject = Nothing
        For Each oMgmtObj In oSearcher.Get()
            Dim installDate As String = oMgmtObj("InstallDate").ToString()
            If Not String.IsNullOrEmpty(installDate) AndAlso installDate.Length >= 14 Then
                Dim year As String = installDate.Substring(0, 4)
                Dim month As String = installDate.Substring(4, 2)
                Dim day As String = installDate.Substring(6, 2)
                Dim hour As String = installDate.Substring(8, 2)
                Dim minute As String = installDate.Substring(10, 2)
                Dim formattedDate As String = $"{year}-{month}-{day} {hour}:{minute}"
                Return formattedDate
            End If
        Next
        Return "Unknown"
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub


End Class