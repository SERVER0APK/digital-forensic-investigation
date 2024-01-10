Imports System.IO
Imports System.Management
Imports Microsoft.Win32
Public Class Form1
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
        Label1.Text = "عدد USB : " & count.ToString()
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


    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub INFO_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
    End Sub

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
        End Using
        MessageBox.Show("Information saved successfully in the file.")
    End Sub


End Class
