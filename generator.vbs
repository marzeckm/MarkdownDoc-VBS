'! The main class of the script
'! 
'! @author { Maximilian Marzeck }
'! @class { Main }
'! 
Class Main
    Dim objFSO, objTranspiler

    '! The main function
    '! 
    '! @public @function main
    '! @return { void }
    '!
    Public Sub main()
        forceConsoleMode()

        ' read the index.md
        Set objFSO = CreateObject("Scripting.FileSystemObject")

        ' copy all files to dist
        If objFSO.FolderExists("./dist/") Then
            DeleteFilesAndFolders "./dist"
        End If
        CopyAllFilesAndFolders "./src", "./dist"

        ' create transpiler object
        Set objTranspiler = new Transpiler
        If objFSO.FileExists("./src/index.md") Then
            PrintInfo("Transpiler is reading file '/src/index.md'.")
            objTranspiler.generateHtml(readFile("./src/index.md"))
        Else
            PrintError("Could not find '/src/index.md'. Please make sure it exists in the corresponding directory.")
        End If

        ' blocks automatic closing
        WScript.echo vbCrLf & "Transpiling successfull. Press Enter to exit..."
        WScript.StdIn.ReadLine()
    End Sub

    '! Reads a requested file and returns it as array of lines
    '! 
    '! @public @function readFile
    '! @param {string} filePath
    '! @return { Array<string> }
    '! 
    Public Function readFile(filePath)
        Dim objFile: Set objFile = objFSO.OpenTextFile(filePath, 1)
        Dim arrFile

        Do Until objFile.AtEndOfStream
            line = objFile.ReadLine

            If IsEmpty(arrFile) Or Not IsArray(arrFile) Then
                ReDim arrFile(0)
            Else
                ReDim Preserve arrFile(UBound(arrFile) + 1)
            End If
            
            arrFile(UBound(arrFile)) = line
        Loop

        objFile.Close()
        readFile = arrFile
    End Function

    '! checks if an even number of chars is contained in a string
    '!
    '! @private @function isEvenNumberOfChar
    '! @param { string } str
    '! @param { string } charToCheck
    '! @return { boolean }
    '!
    Public Function isEvenNumberOfChar(str, charToCheck)
        Dim count, i

        count = 0

        For i = 1 To Len(str)
            If Mid(str, i, 1) = charToCheck Then
                count = count + 1
            End If
        Next

        If count Mod 2 = 0 Then
            isEvenNumberOfChar = True
        Else
            isEvenNumberOfChar = False
        End If
    End Function

    '! Prints an error to console
    '!
    '! @public @function PrintError
    '! @param { string } error
    '! @return { void }
    '!
    Public Sub PrintError(error)
        Wscript.Echo "Error: " & error
    End Sub

    '! prints an info to console
    '!
    '! @public @function PrintInfo
    '! @param { string } info
    '! @return { void }
    '!
    Public Sub PrintInfo(info)
        Wscript.Echo "Info: " & info
    End Sub

    '! Forces the application to run in CScript.exe
    '!
    '! @private @function forceConsoleMode
    '! @return { void }
    '!
    Private Sub forceConsoleMode()
        Dim strArgs, strCmd, strEngine, i, objDebug, wshShell

        Set wshShell = CreateObject( "WScript.Shell" )
        strEngine = UCase( Right( WScript.FullName, 12 ) )

        If strEngine <> "\CSCRIPT.EXE" Then
            strArgs = ""
            
            If WScript.Arguments.Count > 0 Then
                For i = 0 To WScript.Arguments.Count - 1
                    strArgs = strArgs & " " & WScript.Arguments(i)
                Next
            End If

            strCmd = "CSCRIPT.EXE //NoLogo """ & WScript.ScriptFullName & """" & strArgs
            Set objDebug = wshShell.Exec( strCmd )

            Do While objDebug.Status = 0
                WScript.Sleep 100
            Loop

            WScript.Quit objDebug.ExitCode
        End If
    End Sub

    '! Deletes all the old files from /dist
    '! 
    '! @private @function DeleteFilesAndFolders
    '! @param { string } folderPath
    '! @return { void }
    '! 
    Private Sub DeleteFilesAndFolders(folderPath)
        Dim objFolder, objFile, objSubFolder
        
        If objFSO.FolderExists(folderPath) Then
            Set objFolder = objFSO.GetFolder(folderPath)
            
            For Each objFile In objFolder.Files
                If LCase(objFSO.GetFileName(objFile.Path)) <> "index.html" Then
                    objMain.PrintInfo("Deleting file """ & objFile.Path & """")
                    objFSO.DeleteFile(objFile.Path)
                End If
            Next
            
            For Each objSubFolder In objFolder.SubFolders
                objMain.PrintInfo("Deleting folder """ & objSubFolder.Path & """")
                objFSO.DeleteFolder objSubFolder.Path, True
            Next

        End If
    End Sub

    '! Copies all the files from /src to /dist
    '! 
    '! @private @function CopyAllFilesAndFolders
    '! @param { string } folderPath
    '! @param { string } targetPath
    '! @return { void }
    '! 
    Private Sub CopyAllFilesAndFolders(folderPath, targetPath)
        Dim objShell
        Set objShell = CreateObject("WScript.Shell")

        If Not objFSO.FolderExists(targetPath) Then
            objFSO.CreateFolder(targetPath)
            objMain.PrintInfo("Folder was created: " & targetPath)
        End If

        ' copy the whole folder
        objMain.PrintInfo("Copying folder """ & folderPath & """ to """ & targetPath & """")
        objShell.Run ("xcopy """ & folderPath & """ """ & targetPath & """ /E /I /Y"), 1, True

        ' delete index.md, if copied
        If(objFSO.FileExists("./dist/index.md")) Then
            objFSO.DeleteFile("./dist/index.md")
        End If
    End Sub
End Class

'! The transpiler class for creating html from markdown
'! 
'! @author { Maximilian Marzeck }
'! @class { Main }
'! 
Class Transpiler
    Dim objFSO
    Dim dictOrderedLists, dictUnorderedLists
    Dim inCodeBlock, inList
    Dim strCodeBlock, strResult, listKey
    Dim currentListLevel

    '! Mainfunction for the generating of html from markdown
    '!
    '! @public @function generateHtml
    '! @param { Array<string> } fileIndexMd
    '! @return { void }
    '!
    Public Sub generateHtml(fileIndexMd)
        Dim i, line

        Set objFSO = CreateObject("Scripting.FileSystemObject")
        setInitialValues()
       
        ' process other elements
        For i = LBound(fileIndexMd) To UBound(fileIndexMd)
            line = fileIndexMd(i)

            If Trim(line) = "```" Then
                If inCodeBlock Then
                    strResult = strResult & "<pre><code>" & strCodeBlock & "</code></pre>" & vbCrLf
                    setInitialValues()
                Else
                    inCodeBlock = True
                End If
            ElseIf inCodeBlock Then
                strCodeBlock = strCodeBlock & Replace(Replace(line, "<", "&lt;"), ">", "&gt;") & vbCrLf                
            ElseIf checkIfUnorderedList(line) And (Left(line, 1) = "-" Or Left(line, 1) = "*" Or Left(line, 1) = "+") Then
                processListEntry "ul", "UNORDERED_LIST_", line
            ElseIf IsNumeric(Left(line, 1)) And InStr(line, ".") > 0 Then
                processListEntry "ol", "ORDERED_LIST_", line
            Else
                strResult = strResult & processLine(fileIndexMd(i)) & vbCrLf
            End If
        Next

        If inCodeBlock Then
            strResult = strResult & "<pre><code>" & strCodeBlock & "</code></pre>" & vbCrLf
        End If

        ' Ersetze Platzhalter für Listen
        Dim key
        For Each key In dictUnorderedLists.Keys
            strResult = Replace(strResult, "[[UNORDERED_LIST_" & key & "]]", "<ul>" & dictUnorderedLists(key) & "</ul>")
        Next

        For Each key In dictOrderedLists.Keys
            strResult = Replace(strResult, "[[ORDERED_LIST_" & key & "]]", "<ol>" & dictOrderedLists(key) & "</ol>")
        Next

        ' Read the HTML-Template
        Dim strTemplate 
        If objFSO.FileExists("./index.html") Then
            objmain.PrintInfo("Transpiler is reading the template file '/index.html'.")
            strTemplate = Join(objMain.readFile("./index.html"), vbCrLf)

            ' Create the html file of documentation
            writeResultToIndexHtml "./dist/index.html", Replace(strTemplate, "[[CONTENT]]", strResult)
            objMain.printInfo("Index.html was successfully created in folder '/dist/'.")
        Else 
            objMain.printInfo("Could not find '/src/index.md'. Please make sure it exists in the corresponding directory.")
        End If
    End Sub

    '! Mainfunction for the generating of html from markdown
    '!
    '! @private @function setInitialValues
    '! @return { void }
    '!
    Private Sub setInitialValues()
        ' Set Dictionaries
        If IsEmpty(dictOrderedLists) Then
            Set dictOrderedLists = CreateObject("Scripting.Dictionary")
            Set dictUnorderedLists = CreateObject("Scripting.Dictionary")
        End If

        ' Set strings for code block
        inCodeBlock = False
        strCodeBlock = ""

        ' set strings for Lists
        listType = ""
        inList = False

        If IsEmpty(currentListLevel) Then
            currentListLevel = 0
        End if
    End Sub
    
    '! function for processing a line that is 
    '! not part of an multiline element
    '!
    '! @private @function processLine
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processLine(line)
        Dim strResult
        strResult = ""

        setInitialValues()

        ' Remove whitespace and draw lines
        line = Trim(Replace(line, "  ", "<br>"))
        line = Replace(line, "---", "<hr>")

        ' Process other elements
        line = processHeaders(line)
        line = processImages(line)
        line = processLinks(line)
        line = processBoldItalic(line)
        line = processBold(line)
        line = processItalic(line)
        line = processCode(line)
        line = processBlockQuote(line)

        processLine = strResult & line
    End Function

    '! Overrides the index.html file in dist
    '!
    '! @private @function writeResultToIndexHtml
    '! @param { string } filePath
    '! @param { string } line
    '! @return { string }
    '!
    Private Sub writeResultToIndexHtml(filePath, text)
        Set file = objFSO.OpenTextFile(filePath, 2, True)
        file.WriteLine(text)
    End Sub

    '! Processes h1-h6 from markdown to html
    '!
    '! @private @function processHeaders
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processHeaders(line)
        line = Trim(line)
        strResult = line

        If Left(line, 1) = "#" Then
            Dim level, strResult
            ' counst the number of "#"
            level = Len(Split(line, " ", -1, 1)(0))
            line = Trim(Mid(line, level + 1))

            strResult = "<h" & level & ">" & line & "</h" & level & ">"
        End If

        processHeaders = strResult
    End Function

    '! Processes links from markdown to html
    '!
    '! @private @function processLinks
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processLinks(line)
        Do
            Dim textStart, textEnd, urlStart, urlEnd, linkText, linkURL
            
            textStart = InStr(line, "[")
            textEnd = InStr(line, "]")
            urlStart = InStr(line, "(")
            urlEnd = InStr(line, ")")
            
            If textStart >= 0 And textEnd > 0 And urlStart >= 0 And urlEnd > 0 And textEnd < urlStart Then
                ' extract the link data
                linkText = Mid(line, textStart + 1, textEnd - textStart - 1)
                linkURL = Mid(line, urlStart + 1, urlEnd - urlStart - 1)
                
                ' swap the markup with html 
                line = Left(line, textStart - 1) & "<a href=""" & linkURL & """ target=""_blank"">" & linkText & "</a>" & Mid(line, urlEnd + 1)
            Else
                Exit Do
            End If
        Loop
        processLinks = line
    End Function

    '! Processes images from markdown to html
    '!
    '! @private @function processImages
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processImages(line)
        Do
            Dim altTextStart, altTextEnd, urlStart, urlEnd, altText, imgURL
            
            altTextStart = InStr(line, "![")
            altTextEnd = InStr(line, "]")
            urlStart = InStr(line, "(")
            urlEnd = InStr(line, ")")
            
            If altTextStart > 0 And altTextEnd > 0 And urlStart > 0 And urlEnd > 0 And altTextEnd < urlStart Then
                ' extract the alt-text and img url
                altText = Mid(line, altTextStart + 2, altTextEnd - altTextStart - 2)
                imgURL = Mid(line, urlStart + 1, urlEnd - urlStart - 1)
                
                ' swarp the markup ith html
                line = Left(line, altTextStart - 1) & "<div class=""image-wrapper""><img src=""" & imgURL & """ alt=""" & altText & """></div>" & Mid(line, urlEnd + 1)
            Else
                Exit Do
            End If
        Loop
        processImages = line
    End Function

    '! Processes bold-italic elements from markdown to html
    '!
    '! @private @function processBoldItalic
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processBoldItalic(line)
        Do
            Dim boldItalicStart, boldItalicEnd, boldItalicText
            
            boldItalicStart = InStr(line, "***")
            If boldItalicStart > 0 Then
                boldItalicEnd = InStr(boldItalicStart + 3, line, "***")
            Else
                boldItalicEnd = 0
            End If
            
            If boldItalicStart > 0 And boldItalicEnd > 0 Then
                boldItalicText = Mid(line, boldItalicStart + 3, boldItalicEnd - boldItalicStart - 3)
                line = Left(line, boldItalicStart - 1) & "<strong><em>" & boldItalicText & "</em></strong>" & Mid(line, boldItalicEnd + 3)
            Else
                Exit Do
            End If
        Loop
        processBoldItalic = line
    End Function

    '! Processes bold elements from markdown to html
    '!
    '! @private @function processBold
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processBold(line)
        Do
            Dim boldStart, boldEnd, boldText
            
            boldStart = InStr(line, "**")
            If boldStart > 0 Then
                boldEnd = InStr(boldStart + 2, line, "**")
            Else
                boldEnd = 0
            End If
            
            If boldStart > 0 And boldEnd > 0 Then
                boldText = Mid(line, boldStart + 2, boldEnd - boldStart - 2)
                line = Left(line, boldStart - 1) & "<strong>" & boldText & "</strong>" & Mid(line, boldEnd + 2)
            Else
                Exit Do
            End If
        Loop
        processBold = line
    End Function

    '! Processes italic elements from markdown to html
    '!
    '! @private @function processItalic
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processItalic(line)
        Do
            Dim italicStart, italicEnd, italicText
            
            italicStart = InStr(line, "*")
            If italicStart > 0 Then
                italicEnd = InStr(italicStart + 1, line, "*")
            Else
                italicEnd = 0
            End If
            
            If italicStart > 0 And italicEnd > 0 Then
                italicText = Mid(line, italicStart + 1, italicEnd - italicStart - 1)
                line = Left(line, italicStart - 1) & "<em>" & italicText & "</em>" & Mid(line, italicEnd + 1)
            Else
                Exit Do
            End If
        Loop
        processItalic = line
    End Function

    '! Processes a code block from markdown to html
    '!
    '! @private @function processCode
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processCode(line)
        Do
            Dim codeStart, codeEnd, codeText
            
            codeStart = InStr(line, "`")
            If codeStart > 0 Then
                codeEnd = InStr(codeStart + 1, line, "`")
            Else
                codeEnd = 0
            End If
            
            If codeStart > 0 And codeEnd > 0 Then
                codeText = Mid(line, codeStart + 1, codeEnd - codeStart - 1)
                codetext = Replace(codeText, "<", "&lt;")
                codetext = Replace(codeText, ">", "&gt;")
                line = Left(line, codeStart - 1) & "<code class=""inline"">" & codeText & "</code>" & Mid(line, codeEnd + 1)
            Else
                Exit Do
            End If
        Loop
        processCode = line
    End Function

    '! Processes a blockquote from markdown to html
    '!
    '! @private @function processBlockQuote
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processBlockQuote(line)
        If Left(line, 1) = ">" Then
            line = "<blockquote>" & Mid(line, 2) & "</blockquote>"
        End If
        processBlockQuote = line
    End Function

    '! Processes a list entry form markdown to html
    '!
    '! @private @function processListEntry
    '! @param { string } listType
    '! @param { string } placeholder
    '! @param { string } line
    '! @return { string }
    '!
    Private Function processListEntry(listType, placeholder, line)
        If Not inList Then
            currentListLevel = currentListLevel + 1
            listKey = listType & currentListLevel
            dictUnorderedLists.Add listKey, ""
            strResult = strResult & "[[" & placeholder & listKey & "]]" & vbCrLf
            inList = True
        End If

        If listType = "ol" Then
            dictOrderedLists(listKey) = dictOrderedLists(listKey) & "<li>" & Mid(line, InStr(line, ".") + 1) & "</li>" & vbCrLf
        Else 
            dictUnorderedLists(listKey) = dictUnorderedLists(listKey) & "<li>" & Mid(line, 2) & "</li>" & vbCrLf
        End If
    End Function

    '! Checks if a line is within an unordered list
    '!
    '! @private @function checkIfUnorderedList
    '! @param { string } line
    '! @return { string }
    '!
    Private Function checkIfUnorderedList(line)
        line = Trim(line)
        checkIfUnorderedList = True

        If(Left(line, 3) = "***") Then
            checkIfUnorderedList = False
        End if

        If(Left(line, 3) = "---") Then
            checkIfUnorderedList = False
        End if

        If(objMain.isEvenNumberOfChar(line, "*") And objMain.isEvenNumberOfChar(line, "-")) Then
            checkIfUnorderedList = False
        End If
    End Function
End Class

'! The main object is created and main-function started
Dim objMain: Set objMain = new Main
objMain.main()
