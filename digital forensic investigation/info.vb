Imports System.IO
Imports System.Management
Imports Microsoft.Win32
Imports Microsoft.Win32.TaskScheduler

Public Class info
    Private Sub info_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' احصل على معلومات النظام
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        Dim collection As ManagementObjectCollection = searcher.Get()

        ' اطبع معلومات النظام
        For Each info As ManagementObject In collection
            TextBox4.Text = (info("Manufacturer")) & " " & (info("Model"))
            'TextBox4.Text = (info("Model"))
        Next
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

    End Sub
    Sub STARTUP()
        ' تحديد مسار التسجيل لبرامج بدء التشغيل
        Dim keyPath As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

        ' فتح مفتاح التسجيل
        Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(keyPath)

        ' التحقق مما إذا كان المفتاح موجودًا
        If key IsNot Nothing Then
            ' الحصول على أسماء القيم (البرامج)
            Dim valueNames As String() = key.GetValueNames()

            ' عرض معلومات حول كل برنامج
            For Each valueName In valueNames
                ' اسم الملف
                Dim fileName As String = key.GetValue(valueName).ToString()

                ' تاريخ تثبيت البرنامج
                Dim installDate As Date = GetInstallDate(fileName)

                ' إضافة المعلومات إلى RichTextBox بدلاً من إعادة تعيينه
                RichTextBox3.AppendText("File name: " & fileName & vbCrLf)
                RichTextBox3.AppendText("File path: " & GetFilePath(fileName) & vbCrLf)
                RichTextBox3.AppendText("Date of installation: " & installDate.ToString() & vbCrLf)
                RichTextBox3.AppendText("------------------------" & vbCrLf)
            Next
        Else
            ' إذا لم يتم العثور على المفتاح
            RichTextBox3.Text = "لا يمكن العثور على مفتاح التسجيل."
        End If

        ' Console.ReadLine() ' لا داعي لهذا في التطبيق الخاص بك
    End Sub


    ' الوظيفة للحصول على مسار ملف البرنامج
    Function GetFilePath(fileName As String) As String
        Try
            Dim fileInfo As New System.IO.FileInfo(fileName)
            Return fileInfo.FullName
        Catch ex As Exception
            Return "لا يمكن العثور على الملف"
        End Try
    End Function

    ' الوظيفة للحصول على تاريخ تثبيت البرنامج
    Function GetInstallDate(fileName As String) As Date
        Try
            Dim fileInfo As New System.IO.FileInfo(fileName)
            Return fileInfo.CreationTime
        Catch ex As Exception
            Return Date.MinValue
        End Try
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' The file name to save the information
        Dim filePath As String = Path.Combine(Application.StartupPath, "File.txt")
        ' Open the text file for writing
        Using writer As New StreamWriter(filePath, True)
            ' Save information from TextBox1
            writer.WriteLine("Installation Date of the Operating System: " & TextBox1.Text)
            writer.WriteLine("===========")
            ' Save information from TextBox2
            writer.WriteLine("System Type and Kernel: " & TextBox2.Text)
            writer.WriteLine("===========")
            ' Save information from TextBox3
            writer.WriteLine("Computer Name: " & TextBox3.Text)
            writer.WriteLine("===========")
            ' Save information from RichTextBox1
            writer.WriteLine("USB Names:")
            writer.WriteLine(RichTextBox1.Text)
            writer.WriteLine("===========")
            ' Save information from RichTextBox2
            writer.WriteLine("Detected Viruses:")
            writer.WriteLine(RichTextBox2.Text)
            writer.WriteLine("===========")
            writer.WriteLine("StartUP:")
            writer.WriteLine(RichTextBox3.Text)
            writer.WriteLine("===========")
        End Using
        MessageBox.Show("Information saved successfully in the file.")
    End Sub
    Sub DisplayScheduledTasks()
        ' إنشاء مهمة جدولة
        Using ts As TaskService = New TaskService()
            ' الحصول على جميع المهام المجدولة
            Dim tasks As IEnumerable(Of Task) = ts.AllTasks

            ' عرض معلومات حول كل مهمة مجدولة
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
        ' عرض عدد friendlyNames في Label1
        'Label1.Text = " USB : " & count.ToString()
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