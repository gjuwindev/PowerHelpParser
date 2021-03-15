Imports System.Management.Automation

'Install-Package Microsoft.PowerShell.5.1.ReferenceAssemblies [ -Version 1.0.0 ]

Module Module1

    Sub Main()

        Dim command As String = "Get-Help Get-Mailbox -Full"


        Dim result As String = GetPowerShellResult(command)

        ' will process the command output here
        Console.WriteLine(result)


        Console.WriteLine("GOTOVO")
        Console.ReadLine()

    End Sub

    Function GetPowerShellResult(command As String, Optional savedContentOverride As Boolean = False) As String

        Dim powerShellResultsFilename As String = "powerShellResults"

        Dim result As String

        If Not savedContentOverride AndAlso
            IO.File.Exists(powerShellResultsFilename) Then
            result = IO.File.ReadAllText(powerShellResultsFilename)
        Else
            Dim pws As PowerShell = PowerShell.Create()

            result = ExecutePowerShellCommand(pws, command)

            IO.File.WriteAllText(powerShellResultsFilename, result)
        End If

        Return result

    End Function

    Function ExecutePowerShellCommand(pws As PowerShell, command As String) As String

        Dim sb As New Text.StringBuilder

        pws.AddScript(command & " | Out-String")

        Dim psobjects As ObjectModel.Collection(Of PSObject) = pws.Invoke

        For Each pso In psobjects
            sb.AppendLine(pso.ToString)
        Next

        Return sb.ToString

    End Function

End Module
