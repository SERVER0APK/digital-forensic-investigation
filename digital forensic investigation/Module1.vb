Imports System.Management
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Win32


Module Module1

    Sub now()
        Try
            Dim systemInfo As SystemInfo = GetSystemInfo()
            Dim displayInfo As DisplayInfo = GetDisplayInfo()
            Dim processorInfo As ProcessorInfo = GetProcessorInfo()
            Dim hardDriveInfos As List(Of HardDriveInfo) = GetHardDriveInfos()
            Dim biosSerial As String = GetBiosSerial()

            DisplaySystemInfo(systemInfo, processorInfo, hardDriveInfos, biosSerial)
            DisplayDisplayInfo(displayInfo)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Function GetDirectXVersion() As String
        Try
            Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\DirectX")
            If key IsNot Nothing Then
                Dim version As Object = key.GetValue("Version")
                If version IsNot Nothing Then
                    Return version.ToString()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving DirectX version: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return "N/A"
    End Function

    Private Function GetBiosSerial() As String
        Try
            ' Use ManagementObjectSearcher to retrieve BIOS information
            Dim searcher As New ManagementObjectSearcher("root\CIMv2", "SELECT * FROM Win32_BIOS")

            For Each queryObj As ManagementObject In searcher.Get()
                Return queryObj("SerialNumber").ToString()
            Next
        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving BIOS serial number: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return "N/A"
    End Function

    Public Function GetSystemInfo() As SystemInfo
        Dim systemInfo As New SystemInfo()

        Try
            ' System information
            systemInfo.ComputerName = System.Environment.MachineName
            systemInfo.OperatingSystem = $"{System.Environment.OSVersion.Platform} {System.Environment.OSVersion.Version}"
            systemInfo.Language = System.Globalization.CultureInfo.InstalledUICulture.DisplayName
            systemInfo.SystemManufacturer = "N/A" ' Cannot retrieve this information directly
            systemInfo.SystemModel = "N/A" ' Cannot retrieve this information directly
            systemInfo.BIOS = "N/A" ' Cannot retrieve this information directly
            systemInfo.Processor = System.Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER")
            systemInfo.Memory = $"{(My.Computer.Info.TotalPhysicalMemory / (1024 ^ 2)).ToString("N0")} MB RAM"
            systemInfo.PageFile = $"Used: {My.Computer.Info.TotalVirtualMemory - My.Computer.Info.AvailableVirtualMemory} MB, Available: {My.Computer.Info.AvailableVirtualMemory} MB"
            systemInfo.DirectXVersion = GetDirectXVersion()

        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving system information: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return systemInfo
    End Function

    Public Function GetProcessorInfo() As ProcessorInfo
        Dim processorInfo As New ProcessorInfo()

        Try
            ' Use ManagementObjectSearcher to retrieve processor information
            Dim searcher As New ManagementObjectSearcher("root\CIMv2", "SELECT * FROM Win32_Processor")

            For Each queryObj As ManagementObject In searcher.Get()
                processorInfo.Name = queryObj("Name").ToString()
                processorInfo.Manufacturer = queryObj("Manufacturer").ToString()
                processorInfo.Architecture = queryObj("Architecture").ToString()
                processorInfo.Cores = Convert.ToInt32(queryObj("NumberOfCores")).ToString()
                processorInfo.Threads = Convert.ToInt32(queryObj("NumberOfLogicalProcessors")).ToString()
                processorInfo.MaxClockSpeed = (Convert.ToInt32(queryObj("MaxClockSpeed")) / 1000).ToString() & " GHz"
            Next
        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving processor information: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return processorInfo
    End Function

    Public Function GetHardDriveInfos() As List(Of HardDriveInfo)
        Dim hardDriveInfos As New List(Of HardDriveInfo)()

        Try
            ' Use ManagementObjectSearcher to retrieve hard drive information
            Dim searcher As New ManagementObjectSearcher("root\CIMv2", "SELECT * FROM Win32_DiskDrive")

            For Each queryObj As ManagementObject In searcher.Get()
                Dim hardDriveInfo As New HardDriveInfo()
                hardDriveInfo.Name = queryObj("Caption").ToString()
                hardDriveInfo.Manufacturer = queryObj("Manufacturer").ToString()
                hardDriveInfo.InterfaceType = queryObj("InterfaceType").ToString()
                hardDriveInfo.Size = (Convert.ToInt64(queryObj("Size")) / (1024 ^ 3)).ToString("N2") & " GB"
                hardDriveInfos.Add(hardDriveInfo)
            Next
        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving hard drive information: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return hardDriveInfos
    End Function

    Public Function GetDisplayInfo() As DisplayInfo
        Dim displayInfo As New DisplayInfo()

        Try
            ' Use ManagementObjectSearcher to retrieve display information
            Dim searcher As New ManagementObjectSearcher("root\CIMv2", "SELECT * FROM Win32_VideoController")

            For Each queryObj As ManagementObject In searcher.Get()
                displayInfo.Name = queryObj("Caption").ToString()
                displayInfo.Manufacturer = queryObj("AdapterCompatibility").ToString()
                displayInfo.ChipType = queryObj("VideoProcessor").ToString()
                displayInfo.DacType = "Integrated RAMDAC" ' Constant value
                displayInfo.DeviceType = "Full Display Device" ' Constant value
                displayInfo.ApproxTotalMemory = (Convert.ToInt64(queryObj("AdapterRAM")) / (1024 ^ 2)).ToString("N0") & " MB"
                displayInfo.DisplayMemory = (Convert.ToInt64(queryObj("AdapterRAM")) / (1024 ^ 2)).ToString("N0") & " MB" ' Constant value
                displayInfo.SharedMemory = "N/A" ' Cannot retrieve this information directly
                displayInfo.CurrentDisplayMode = $"{queryObj("CurrentHorizontalResolution")} x {queryObj("CurrentVerticalResolution")} ({queryObj("CurrentBitsPerPixel")} bit) ({queryObj("CurrentRefreshRate")}Hz)"
                displayInfo.Monitor = "N/A" ' Cannot retrieve this information directly
                displayInfo.Hdr = "N/A" ' Constant value
            Next
        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving display information: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return displayInfo
    End Function

    Private Sub DisplaySystemInfo(systemInfo As SystemInfo, processorInfo As ProcessorInfo, hardDriveInfos As List(Of HardDriveInfo), biosSerial As String)
        Dim infoStringBuilder As New StringBuilder()

        Try
            infoStringBuilder.AppendLine($"Current Date/Time: {DateTime.Now.ToString("dd MMMM, yyyy, hh:mm:ss tt")}")
            infoStringBuilder.AppendLine($"Computer Name: {systemInfo.ComputerName}")
            infoStringBuilder.AppendLine($"Operating System: {systemInfo.OperatingSystem}")
            infoStringBuilder.AppendLine($"Language: {systemInfo.Language}")
            infoStringBuilder.AppendLine($"System Manufacturer: {systemInfo.SystemManufacturer}")
            infoStringBuilder.AppendLine($"System Model: {systemInfo.SystemModel}")
            infoStringBuilder.AppendLine($"BIOS: {systemInfo.BIOS} (Serial: {biosSerial})")
            infoStringBuilder.AppendLine($"Processor: {processorInfo.Name}")
            infoStringBuilder.AppendLine($"Processor Manufacturer: {processorInfo.Manufacturer}")
            infoStringBuilder.AppendLine($"Processor Architecture: {processorInfo.Architecture}")
            infoStringBuilder.AppendLine($"Processor Cores: {processorInfo.Cores}")
            infoStringBuilder.AppendLine($"Processor Threads: {processorInfo.Threads}")
            infoStringBuilder.AppendLine($"Max Clock Speed: {processorInfo.MaxClockSpeed}")
            infoStringBuilder.AppendLine($"Memory: {systemInfo.Memory}")
            infoStringBuilder.AppendLine($"Page file: {systemInfo.PageFile}")

            ' Display hard drive information
            infoStringBuilder.AppendLine($"Hard Drives:")
            For Each hardDriveInfo In hardDriveInfos
                infoStringBuilder.AppendLine($"- Name: {hardDriveInfo.Name}")
                infoStringBuilder.AppendLine($"  Manufacturer: {hardDriveInfo.Manufacturer}")
                infoStringBuilder.AppendLine($"  Interface Type: {hardDriveInfo.InterfaceType}")
                infoStringBuilder.AppendLine($"  Size: {hardDriveInfo.Size}")
            Next

            infoStringBuilder.AppendLine($"DirectX Version: {systemInfo.DirectXVersion}")
            infoStringBuilder.AppendLine(GetOperatingSystemInfo()) ' إضافة معلومات نظام التشغيل

            info.RichTextBox6.Text = infoStringBuilder.ToString()
        Catch ex As Exception
            MessageBox.Show($"An error occurred while displaying system information: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DisplayDisplayInfo(displayInfo As DisplayInfo)
        Dim infoStringBuilder As New StringBuilder()

        Try
            infoStringBuilder.AppendLine($"Name: {displayInfo.Name}")
            infoStringBuilder.AppendLine($"Manufacturer: {displayInfo.Manufacturer}")
            infoStringBuilder.AppendLine($"Chip Type: {displayInfo.ChipType}")
            infoStringBuilder.AppendLine($"DAC Type: {displayInfo.DacType}")
            infoStringBuilder.AppendLine($"Device Type: {displayInfo.DeviceType}")
            infoStringBuilder.AppendLine($"Approx. Total Memory: {displayInfo.ApproxTotalMemory}")
            infoStringBuilder.AppendLine($"Display Memory (VRAM): {displayInfo.DisplayMemory}")
            infoStringBuilder.AppendLine($"Shared Memory: {displayInfo.SharedMemory}")
            infoStringBuilder.AppendLine($"Current Display Mode: {displayInfo.CurrentDisplayMode}")
            infoStringBuilder.AppendLine($"Monitor: {displayInfo.Monitor}")
            infoStringBuilder.AppendLine($"HDR: {displayInfo.Hdr}")

            info.RichTextBox7.Text = infoStringBuilder.ToString()
        Catch ex As Exception
            MessageBox.Show($"An error occurred while displaying display information: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetOperatingSystemInfo() As String
        Try
            Dim searcher As New ManagementObjectSearcher("root\CIMv2", "SELECT * FROM Win32_OperatingSystem")

            For Each queryObj As ManagementObject In searcher.Get()
                Dim caption As String = queryObj("Caption").ToString()
                Dim installDate As String = queryObj("InstallDate").ToString()

                ' تحويل تاريخ التثبيت إلى تنسيق قابل للقراءة
                Dim formattedInstallDate As String = ""
                If installDate.Length >= 8 Then
                    formattedInstallDate = $"{installDate.Substring(6, 2)}/{installDate.Substring(4, 2)}/{installDate.Substring(0, 4)}"
                End If

                Return $"Operating System: {caption} ({formattedInstallDate})"
            Next
        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving operating system information: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return "Operating System: N/A"
    End Function
    'End Class

    Public Class SystemInfo
        Public Property ComputerName As String
        Public Property OperatingSystem As String
        Public Property Language As String
        Public Property SystemManufacturer As String
        Public Property SystemModel As String
        Public Property BIOS As String
        Public Property Processor As String
        Public Property Memory As String
        Public Property PageFile As String
        Public Property DirectXVersion As String
    End Class

    Public Class ProcessorInfo
        Public Property Name As String
        Public Property Manufacturer As String
        Public Property Architecture As String
        Public Property Cores As String
        Public Property Threads As String
        Public Property MaxClockSpeed As String
    End Class

    Public Class HardDriveInfo
        Public Property Name As String
        Public Property Manufacturer As String
        Public Property InterfaceType As String
        Public Property Size As String
    End Class

    Public Class DisplayInfo
        Public Property Name As String
        Public Property Manufacturer As String
        Public Property ChipType As String
        Public Property DacType As String
        Public Property DeviceType As String
        Public Property ApproxTotalMemory As String
        Public Property DisplayMemory As String
        Public Property SharedMemory As String
        Public Property CurrentDisplayMode As String
        Public Property Monitor As String
        Public Property Hdr As String
    End Class
End Module
