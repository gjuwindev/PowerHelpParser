Imports System.Management.Automation

'Install-Package Microsoft.PowerShell.5.1.ReferenceAssemblies [ -Version 1.0.0 ]

'TODO: get list of all commands:  "Get-Command | Select-Object Name" and process help texts for all that commands...

Module Module1

    Class HelpSection
        Public Name As String
        Public Lines As New List(Of String)
    End Class

    Class CommandHelpSections
        Public Name As HelpSection = Nothing
        Public Synopsis As HelpSection = Nothing
        Public Syntax As HelpSection = Nothing
        Public Description As HelpSection = Nothing
        Public Parameters As HelpSection = Nothing
        Public Inputs As HelpSection = Nothing
        Public Outputs As HelpSection = Nothing
        Public Notes As HelpSection = Nothing
        Public RelatedLinks As HelpSection = Nothing
        Public OtherSections As New List(Of HelpSection)
    End Class

    Sub Main()

        Dim command As String = "Get-Help Get-Mailbox -Full"
        Dim helpObject As New CommandHelpSections
        Dim result As String = GetPowerShellResult(command, savedContentOverride:=False)

        ParseResultIntoHelpSections(result, helpObject)
        DistibuteSections(helpObject)
        If helpObject.OtherSections.Count > 0 Then Diagnostics.Debugger.Break()

        PrintCommandHelpContents(helpObject)

        Console.WriteLine("FINISHED")
        Console.ReadLine()

    End Sub

    Sub PrintCommandHelpContents(helpobject As CommandHelpSections)

        Console.WriteLine(helpobject.Name.Name & ": " & helpobject.Name.Lines.Count)
        Console.WriteLine(helpobject.Synopsis.Name & ": " & helpobject.Synopsis.Lines.Count)
        Console.WriteLine(helpobject.Syntax.Name & ": " & helpobject.Syntax.Lines.Count)
        Console.WriteLine(helpobject.Description.Name & ": " & helpobject.Description.Lines.Count)
        Console.WriteLine(helpobject.Parameters.Name & ": " & helpobject.Parameters.Lines.Count)
        Console.WriteLine(helpobject.Inputs.Name & ": " & helpobject.Inputs.Lines.Count)
        Console.WriteLine(helpobject.Outputs.Name & ": " & helpobject.Outputs.Lines.Count)
        Console.WriteLine(helpobject.Notes.Name & ": " & helpobject.Notes.Lines.Count)
        Console.WriteLine(helpobject.RelatedLinks.Name & ": " & helpobject.RelatedLinks.Lines.Count)

        For Each section In helpobject.OtherSections
            Console.WriteLine(section.Name & ": " & section.Lines.Count & " lines")

            Dim counter = 0
            For Each line In section.Lines
                counter += 1
                If counter > 5 Then Exit For
                Console.WriteLine("  " & line)
            Next
        Next

    End Sub

    Private Sub DistibuteSections(helpObject As CommandHelpSections)

        Dim numberOfSections As Integer = helpObject.OtherSections.Count

        For n = numberOfSections - 1 To 0 Step -1
            Dim section As HelpSection = helpObject.OtherSections(n)
            Dim name = section.Name.ToUpper

            Dim found As Boolean = True
            If helpObject.Name Is Nothing AndAlso name = "NAME" Then
                helpObject.Name = section
            ElseIf helpObject.Synopsis Is Nothing AndAlso name = "SYNOPSIS" Then
                helpObject.Synopsis = section
            ElseIf helpObject.Syntax Is Nothing AndAlso name = "SYNTAX" Then
                helpObject.Syntax = section
            ElseIf helpObject.Description Is Nothing AndAlso name = "DESCRIPTION" Then
                helpObject.Description = section
            ElseIf helpObject.Parameters Is Nothing AndAlso name = "PARAMETERS" Then
                helpObject.Parameters = section
            ElseIf helpObject.Inputs Is Nothing AndAlso name = "INPUTS" Then
                helpObject.Inputs = section
            ElseIf helpObject.Outputs Is Nothing AndAlso name = "OUTPUTS" Then
                helpObject.Outputs = section
            ElseIf helpObject.Notes Is Nothing AndAlso name = "NOTES" Then
                helpObject.Notes = section
            ElseIf helpObject.RelatedLinks Is Nothing AndAlso name = "RELATED LINKS" Then
                helpObject.RelatedLinks = section
            Else
                found = False
            End If

            If found Then
                helpObject.OtherSections.RemoveAt(n)
            End If
        Next

    End Sub

    Private Sub ParseResultIntoHelpSections(result As String, helpObject As CommandHelpSections)

        Dim lines() As String = result.Split({vbCrLf}, StringSplitOptions.None)

        Dim sectionLines As New List(Of String)
        Dim section As HelpSection = Nothing

        For Each line In lines
            If line.Trim = "" OrElse line.StartsWith(" ") Then
                If section Is Nothing Then
                    'Diagnostics.Debugger.Break()
                Else
                    sectionLines.Add(line)
                End If
            Else ' novi naslov sekcije
                If section IsNot Nothing Then
                    section.Lines = sectionLines
                End If

                section = New HelpSection
                section.Name = line.Trim
                sectionLines = New List(Of String)
                sectionLines.Append(line)
                helpObject.OtherSections.Add(section)
            End If
        Next

        For Each section In helpObject.OtherSections

            While section.Lines.Count > 0
                If section.Lines(0).Trim = "" Then
                    section.Lines.RemoveAt(0)
                Else
                    Exit While
                End If
            End While

            Dim lastLineNumber As Integer = section.Lines.Count - 1
            While lastLineNumber > -1
                If section.Lines(lastLineNumber).Trim = "" Then
                    section.Lines.RemoveAt(lastLineNumber)
                    lastLineNumber -= 1
                Else
                    Exit While
                End If
            End While

        Next

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
