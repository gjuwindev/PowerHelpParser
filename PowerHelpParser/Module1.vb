Imports System.Management.Automation

'Install-Package Microsoft.PowerShell.5.1.ReferenceAssemblies [ -Version 1.0.0 ]

'TODO: get list of all commands: "Get-Command | Select-Object Name" and process help texts for all that commands...

Module Module1

    Class HelpSection
        Public Name As String
        Public Lines As New List(Of String)
    End Class

    Class Parameter
        Public Name As String
        Public Type As String
        Public Description As New List(Of String)

        Public Required As String
        Public Position As String
        Public DefaultValue As String
        Public AcceptPipelineInput As String
        Public AcceptWildcardCharacters As String

        Public OtherLines As New List(Of String)
    End Class

    Class CommandHelpSections
        Public NameSection As HelpSection = Nothing
        Public SynopsisSection As HelpSection = Nothing
        Public SyntaxSection As HelpSection = Nothing
        Public DescriptionSection As HelpSection = Nothing
        Public ParametersSection As HelpSection = Nothing
        Public Parameters As New List(Of Parameter)
        Public InputsSection As HelpSection = Nothing
        Public OutputsSection As HelpSection = Nothing
        Public NotesSection As HelpSection = Nothing
        Public RelatedLinksSection As HelpSection = Nothing
        Public OtherSections As New List(Of HelpSection)
    End Class

    Sub Main()

        Dim command As String = "Get-Help Get-Mailbox -Full"
        Dim helpObject As New CommandHelpSections
        Dim result As String = GetPowerShellResult(command, savedContentOverride:=False)

        ParseResultIntoHelpSections(result, helpObject)
        DistibuteSections(helpObject)
        If helpObject.OtherSections.Count > 0 Then Diagnostics.Debugger.Break()
        AnalyzeParameters(helpObject)

        PrintCommandHelpContents(helpObject)

        Console.WriteLine("FINISHED")
        Console.ReadLine()

    End Sub

    Private Sub AnalyzeParameters(helpObject As CommandHelpSections)

        SeparateParameterTextIntoSeparateSetsOfLines(helpObject)

        Dim trimmedLine As String

        For Each parm In helpObject.Parameters
            trimmedLine = parm.OtherLines(0).Trim
            If trimmedLine = "<CommonParameters>" Then
                parm.Name = trimmedLine
                parm.Type = "<various>"
            Else
                Dim blankPos As Integer = trimmedLine.IndexOf(" ")
                If blankPos = -1 Then System.Diagnostics.Debugger.Break()

                parm.Name = trimmedLine.Substring(1, blankPos - 1)
                parm.Type = trimmedLine.Substring(blankPos + 1).TrimStart("<").TrimEnd(">")
            End If

            parm.OtherLines.RemoveAt(0)

            Dim lineCount As Integer = parm.OtherLines.Count

            For n = lineCount - 1 To 0 Step -1
                trimmedLine = parm.OtherLines(n).Trim
                Dim found = True
                If trimmedLine = "" Then
                    '
                ElseIf trimmedLine.StartsWith("Required") Then
                    parm.Required = trimmedLine.Substring(9).Trim
                ElseIf trimmedLine.StartsWith("Position") Then
                    parm.Position = trimmedLine.Substring(9).Trim
                ElseIf trimmedLine.StartsWith("Default value") Then
                    parm.DefaultValue = trimmedLine.Substring(13).Trim
                ElseIf trimmedLine.StartsWith("Accept pipeline input") Then
                    parm.AcceptPipelineInput = trimmedLine.Substring(22).Trim
                ElseIf trimmedLine.StartsWith("Accept wildcard characters") Then
                    parm.AcceptWildcardCharacters = trimmedLine.Substring(27).Trim
                Else
                    found = False
                End If

                If found Then
                    parm.OtherLines.RemoveAt(n)
                Else
                    Exit For
                End If
            Next

        Next

    End Sub

    Private Sub SeparateParameterTextIntoSeparateSetsOfLines(helpObject As CommandHelpSections)

        Dim parm As Parameter
        Dim lineList As List(Of String) = Nothing

        For Each line As String In helpObject.ParametersSection.Lines
            Dim trimmedLine As String = line.Trim
            If (trimmedLine.StartsWith("-") AndAlso Not trimmedLine.StartsWith("- ")) OrElse trimmedLine = "<CommonParameters>" Then
                If lineList IsNot Nothing Then
                    parm = New Parameter
                    parm.OtherLines = lineList
                    helpObject.Parameters.Add(parm)
                End If
                parm = New Parameter
                lineList = New List(Of String)
                parm.OtherLines = lineList
                lineList.Add(line)
            ElseIf lineList Is Nothing Then
                System.Diagnostics.Debugger.Break()
            Else
                lineList.Add(line)
            End If
        Next

    End Sub

    Sub PrintCommandHelpContents(helpobject As CommandHelpSections)

        Console.WriteLine(helpobject.NameSection.Name & ":  " & helpobject.NameSection.Lines.Count)
        Console.WriteLine(helpobject.SynopsisSection.Name & ":  " & helpobject.SynopsisSection.Lines.Count)
        Console.WriteLine(helpobject.SyntaxSection.Name & ":  " & helpobject.SyntaxSection.Lines.Count)
        Console.WriteLine(helpobject.DescriptionSection.Name & ":  " & helpobject.DescriptionSection.Lines.Count)
        Console.WriteLine(helpobject.ParametersSection.Name & ":  " & helpobject.ParametersSection.Lines.Count)
        PrintParameters(helpobject.Parameters)
        Console.WriteLine(helpobject.InputsSection.Name & ":  " & helpobject.InputsSection.Lines.Count)
        Console.WriteLine(helpobject.OutputsSection.Name & ":  " & helpobject.OutputsSection.Lines.Count)
        Console.WriteLine(helpobject.NotesSection.Name & ":  " & helpobject.NotesSection.Lines.Count)
        Console.WriteLine(helpobject.RelatedLinksSection.Name & ":  " & helpobject.RelatedLinksSection.Lines.Count)

        For Each section In helpobject.OtherSections
            Console.WriteLine(section.Name & ":" & section.Lines.Count & " lines")

            Dim counter = 0
            For Each line In section.Lines
                counter += 1
                If counter > 5 Then Exit For
                Console.WriteLine(" " & line)
            Next
        Next

    End Sub

    Private Sub PrintParameters(parameters As List(Of Parameter))

        Dim showParameterOptions As Boolean = False

        For Each parm In parameters
            Console.WriteLine("   -" & parm.Name & ":  " & parm.Type)

            If showParameterOptions Then
                Console.WriteLine("         Required?                 " & parm.Required)
                Console.WriteLine("         Position?                 " & parm.Position)
                Console.WriteLine("         DefaultValue:             " & parm.DefaultValue)
                Console.WriteLine("         AcceptPipelineInput?      " & parm.AcceptPipelineInput)
                Console.WriteLine("         AcceptWildcardCharacters? " & parm.AcceptWildcardCharacters)
            End If
        Next
    End Sub

    Private Sub DistibuteSections(helpObject As CommandHelpSections)

        Dim numberOfSections As Integer = helpObject.OtherSections.Count

        For n = numberOfSections - 1 To 0 Step -1
            Dim section As HelpSection = helpObject.OtherSections(n)
            Dim name = section.Name.ToUpper

            Dim found As Boolean = True
            If helpObject.NameSection Is Nothing AndAlso name = "NAME" Then
                helpObject.NameSection = section
            ElseIf helpObject.SynopsisSection Is Nothing AndAlso name = "SYNOPSIS" Then
                helpObject.SynopsisSection = section
            ElseIf helpObject.SyntaxSection Is Nothing AndAlso name = "SYNTAX" Then
                helpObject.SyntaxSection = section
            ElseIf helpObject.DescriptionSection Is Nothing AndAlso name = "DESCRIPTION" Then
                helpObject.DescriptionSection = section
            ElseIf helpObject.ParametersSection Is Nothing AndAlso name = "PARAMETERS" Then
                helpObject.ParametersSection = section
            ElseIf helpObject.InputsSection Is Nothing AndAlso name = "INPUTS" Then
                helpObject.InputsSection = section
            ElseIf helpObject.OutputsSection Is Nothing AndAlso name = "OUTPUTS" Then
                helpObject.OutputsSection = section
            ElseIf helpObject.NotesSection Is Nothing AndAlso name = "NOTES" Then
                helpObject.NotesSection = section
            ElseIf helpObject.RelatedLinksSection Is Nothing AndAlso name = "RELATED LINKS" Then
                helpObject.RelatedLinksSection = section
            Else
                found = False
            End If

            If found Then
                helpObject.OtherSections.RemoveAt(n)
            End If
        Next

    End Sub

    Private Sub ParseResultIntoHelpSections(result As String, helpObject As CommandHelpSections)

        ' break help text into lines
        Dim lines() As String = result.Split({vbCrLf}, StringSplitOptions.None)

        ' lines for one section
        Dim sectionLines As List(Of String) = Nothing
        ' one section
        Dim section As HelpSection = Nothing

        For Each line In lines
            If line.StartsWith(" ") OrElse line.Trim = "" Then
                If section Is Nothing Then
                    'Diagnostics.Debugger.Break()
                ElseIf sectionLines Is Nothing And line.Trim = "" Then
                    ' ignore leading empty lines
                Else
                    If sectionLines Is Nothing Then sectionLines = New List(Of String)
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
