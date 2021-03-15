Imports System.Management.Automation

'Install-Package Microsoft.PowerShell.5.1.ReferenceAssemblies [ -Version 1.0.0 ]

Module Module1

    Sub Main()

        Dim pws = PowerShell.Create()
        pws.AddScript("Get-Help Get-Mailbox -Full | Out-String")

        Dim psobjects As ObjectModel.Collection(Of PSObject) = pws.Invoke

        For Each pso In psobjects
            Console.WriteLine(pso)
        Next

        Console.WriteLine("GOTOVO")
        Console.ReadLine()

    End Sub

End Module
