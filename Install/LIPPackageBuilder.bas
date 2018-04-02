Attribute VB_Name = "LIPPackageBuilder"
Option Explicit

' Used for showing in the GUI and also setting in the generated packages.json files.
Private Const m_sLIPPackageBuilderVersion As String = "1.2.0"

' ##SUMMARY Opens the LBS app in a HTML window.
Public Sub OpenPackageBuilder()
    On Error GoTo ErrorHandler
    
    Dim oDialog As New Lime.Dialog
    Dim idpersons As String
    Dim oItem As Lime.ExplorerItem
    oDialog.Type = lkDialogHTML
    oDialog.Property("url") = Application.WebFolder & "lbs.html?ap=apps/LIPPackageBuilder/packagebuilder&type=tab"
    oDialog.Property("height") = 900
    oDialog.Property("width") = 1600
    Call oDialog.show

    Exit Sub
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.OpenPackageBuilder")
End Sub


' ##SUMMARY Called from javascript.
Public Function LoadDataStructure(strProcedureName As String) As String
    On Error GoTo ErrorHandler
    
    Dim oProcedure As LDE.Procedure
    Dim sXml As String
    Set oProcedure = Database.Procedures.Lookup(strProcedureName, lkLookupProcedureByName)
    If Not oProcedure Is Nothing Then
        oProcedure.Parameters("@@lang").InputValue = Database.Locale
        oProcedure.Parameters("@@idcoworker").InputValue = ActiveUser.Record.ID
        Call oProcedure.Execute(False)
    Else
        Call Application.MessageBox("The procedure """ & strProcedureName & """ does not exist in the client metadata.")
    End If
    sXml = oProcedure.result
    sXml = XMLEncodeBase64(sXml)
    
    LoadDataStructure = sXml
    'MsgBox sXml
    'MsgBox StrConv(DecodeBase64(sXml), vbUnicode)
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.LoadDatastructure")
End Function


' ##SUMMARY Called from javascript.
' Returns a JSON string with an array of objects with two parameters:
' the file name and the table name of all the available Actionpad HTML files in the current solution.
Public Function GetAvailableActionpads() As String
    On Error GoTo ErrorHandler
    
    ' Build array of objects
    Dim sActionpadsJson As String
    sActionpadsJson = "["
    
    Dim sFileName As String
    Dim sClassName As String
    Dim lngCount As Long
    lngCount = 0
    sFileName = Dir(Application.WebFolder & "\*.htm*")
    Do While Len(sFileName) > 0
        lngCount = lngCount + 1
        sClassName = VBA.Left(sFileName, VBA.InStr(sFileName, ".") - 1)
        If Database.Classes.Exists(sClassName) Then
            ' Add comma if there already is an element in the array
            If lngCount > 1 Then
                sActionpadsJson = sActionpadsJson + ","
            End If
            ' Add actionpad as object in the array
            sActionpadsJson = sActionpadsJson & "{""tableName"": """ & sClassName & """, ""fileName"": """ & sFileName & """}"
        End If
        
        sFileName = Dir
    Loop
    
    ' Close array
    sActionpadsJson = sActionpadsJson & "]"
    
    GetAvailableActionpads = sActionpadsJson

    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.GetAvailableActionpads")
End Function


' ##SUMMARY Called from javascript.
Public Function GetVBAComponents() As String
    On Error GoTo ErrorHandler
    
    Dim oComp As Object
    Dim strComponents As String
    strComponents = "["
    For Each oComp In Application.VBE.ActiveVBProject.VBComponents
        'Only include modules, class modules and forms
        If oComp.Type <> 11 And oComp.Type <> 100 Then
            strComponents = strComponents & "{"
            strComponents = strComponents & """name"": """ & oComp.Name & ""","
            strComponents = strComponents & """type"": """ & GetModuleTypeName(oComp.Type) & """},"
        End If
    Next
    
    strComponents = VBA.Left(strComponents, Len(strComponents) - 1)
    strComponents = strComponents + "]"
    
    GetVBAComponents = strComponents
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.GetVBAComponents")
End Function


Private Function GetModuleTypeName(ModuleType As Long) As String
    On Error GoTo ErrorHandler
    
    Dim strModuleTypeName As String
    strModuleTypeName = ""
    Select Case ModuleType
        Case 1:
            strModuleTypeName = "Module"
        Case 2:
            strModuleTypeName = "Class Module"
        Case 3:
            strModuleTypeName = "Form"
        Case Else
            strModuleTypeName = "Other"
    End Select
    
    GetModuleTypeName = strModuleTypeName
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.GetModuleTypeName")
End Function


Private Function XMLEncodeBase64(text As String) As String

    If text = "" Then XMLEncodeBase64 = "": Exit Function
     
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
     
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
     
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
     
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    XMLEncodeBase64 = objNode.text
     
    Set objNode = Nothing
    Set objXML = Nothing
     
End Function


Private Function DecodeBase64(ByVal strData As String) As Byte()
 
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    
    ' help from MSXML
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.text = strData
    DecodeBase64 = objNode.nodeTypedValue
    
    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing
 
End Function


' ##SUMMARY Called from javascript when the user clicks the button to create the package.
' Parameters should be submitted as Base64 encoded JSON strings.
Public Sub CreatePackage(sPackage As String, sMetaData As String, sReadmeInfo As String, sChangelogInfo As String, isAddon As Boolean)
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    bResult = True
    
    ' Set the folder where we can operate
    Dim sGeneralWorkingFolderPath As String
    sGeneralWorkingFolderPath = Application.TemporaryFolder & "\LIPPackageBuilder\"
    
    ' Create a temporary working folder for this package
    Dim sTemporaryPackageFolderPath As String
    sTemporaryPackageFolderPath = CreateFolder(sGeneralWorkingFolderPath, VBA.Replace(VBA.Replace(LCO.GenerateGUID, "{", ""), "}", ""))
    
    ' Create JSON objects from base64 encoded JSON strings in the input parameters.
    Dim oPackage As Object
    Set oPackage = Base64StringToJsonObject(sPackage)
    
    Dim oMetaData As Object
    Set oMetaData = Base64StringToJsonObject(sMetaData)
    
    Dim oReadmeInfo As Object
    Set oReadmeInfo = Base64StringToJsonObject(sReadmeInfo)
    
    Dim oChangelogInfo As Object
    Set oChangelogInfo = Base64StringToJsonObject(sChangelogInfo)
    
    'Export VBA modules
    If bResult And oPackage("install").Exists("vba") Then
        bResult = ExportVBA(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export VBA Modules.", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    
    'Export SQL Procedures and functions
    If bResult And oPackage("install").Exists("sql") Then
        bResult = ExportSql(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export SQL Procedures and functions", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    
    'Export Table icons
    If bResult Then
        bResult = SaveTableIcons(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export table icons, will continue anyway...", vbInformation)
        bResult = True
    End If
    
    'Export option queries
    If bResult Then
        bResult = SaveOptionQueries(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export the optionqueries, will continue anyway...", vbInformation)
        bResult = True
    End If
    
    ' Save SQL on update
    If bResult Then
        bResult = SaveSqlOnUpdate(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export the sql on update, will continue anyway...", vbInformation)
        bResult = True
    End If
    
    ' Save SQL for new
     If bResult Then
        bResult = SaveSqlForNew(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export the sql for new, will continue anyway...", vbInformation)
        bResult = True
    End If
    
    ' Save SQL Expressions
    If bResult Then
        bResult = SaveSqlExpressions(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export the sql expressions, will continue anyway...", vbInformation)
        bResult = True
    End If
    
    ' Save SQL Descriptive
    If bResult Then
        bResult = SaveSqlDescriptive(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export the sql descriptive expressions, will continue anyway...", vbInformation)
        bResult = True
    End If
    
    ' Export Actionpads
    If bResult And oPackage("install").Exists("actionpads") Then
        bResult = ExportActionpads(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("Could not export Actionpads.", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    

    'LJE This is not yet implemented
    'If bResult Then
    '    bResult = CleanupPackageFile(oPackage)
    'End If
    'If Not bResult Then
    '    Call Application.MessageBox("Couldn't cleanup the package file, aborting...", vbError)
    '    bResult = False
    'End If

    
    ' Save Package.json
    If bResult Then
        bResult = SavePackageFile(oPackage, sTemporaryPackageFolderPath)
    End If
    If Not bResult Then
        Call Application.MessageBox("An error occurred: Could not save the package.json file.", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    
    ' ##TODO: Remove if not needed anymore!
    ' Rename Temporary folder to package name
'    Dim sNewFolderPath As String
'    If bResult Then
'        sNewFolderPath = RenameTemporaryFolder(oMetaData.Item("uniqueName"), sTemporaryPackageFolderPath)
'    End If
'
'    If sNewFolderPath = "" Then
'        Call Application.MessageBox("An error occurred: Could not rename the temporary folder.", VBA.vbCritical + VBA.vbOKOnly)
'        Exit Sub
'    End If

    ' Let the user select a folder to place the generated files in
    Dim sSelectedPath As String
    sSelectedPath = GetFolder("Select a folder to save the generated files in.")
    If sSelectedPath = "" Then
        ' User aborted
        Exit Sub
    End If
    
    ' Create folder where to put all generated files (zip file and subfolder for all add-on files).
    Dim sTargetPath As String
    sTargetPath = CreateFolder(sSelectedPath, oMetaData.Item("uniqueName") & "_" & GetCleanTimestamp)
    
    
    
    ' Create zip for LIP Package
    Dim sZipFileFullPath As String
    If bResult Then
        Dim sZipName As String
        
        If isAddon Then
            sZipName = "lip-add-on-" & oMetaData.Item("uniqueName") & "-v" & oChangelogInfo.Item("versionNumber")
        Else
            sZipName = "lip-" & oMetaData.Item("uniqueName")
        End If
        sZipFileFullPath = ZipFolder(sZipName, sTemporaryPackageFolderPath, sTargetPath)
    End If
    
    If sZipFileFullPath = "" Then
        Call Application.MessageBox("An error occurred: Could not save the package zip file.", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    
    ' Save all files under the add-on folder for easy access if it is a new add-on
    Dim sAddonFolderPath As String
    sAddonFolderPath = sTargetPath & "\add-on"
    
    ' Copy generated lip files from the temporary folder to the add-on\lip folder
    Call CreateFolder(sAddonFolderPath, "lip")
    If Not CopyFolder(sTemporaryPackageFolderPath, sAddonFolderPath & "\lip") Then
        Call Application.MessageBox("An error occurred: Could not copy add-on files from temporary folder to target folder.", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    
    ' Create other mandatory files for add-ons
    If Not SaveTextToDisk(VBA.StrConv(DecodeBase64(sMetaData), VBA.vbUnicode), sAddonFolderPath & "\resources", "metadata.json") Then
        Call Application.MessageBox("An error occurred: Could not create metadata.json.", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    
    If Not CreateReadmeMd(oReadmeInfo, sAddonFolderPath) Then
        Call Application.MessageBox("An error occurred: Could not create README.md.", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    
    If Not CreateChangelogMd(oChangelogInfo, sAddonFolderPath) Then
        Call Application.MessageBox("An error occurred: Could not create CHANGELOG.md.", VBA.vbCritical + VBA.vbOKOnly)
        Exit Sub
    End If
    
    ' Notify the user that all went well
    Call Application.MessageBox("All generated files were saved successfully in folder " & sTargetPath & ".", VBA.vbInformation + VBA.vbOKOnly)
    
    ' Open the folder containing the zip file
    Dim sZipFileFolderPath As String
    sZipFileFolderPath = VBA.Left(sZipFileFullPath, VBA.InStrRev(sZipFileFullPath, "\") - 1)
    Call Application.Shell(sZipFileFolderPath)
    
    'Delete Temporary folder
    If bResult Then
        bResult = DeleteTemporaryFolder(sGeneralWorkingFolderPath)
    End If
    
    If Not bResult Then
        Call Application.MessageBox("An error occurred: Could not remove the temporary folder %1", VBA.vbCritical + VBA.vbOKOnly, sGeneralWorkingFolderPath)
    End If
    
    Exit Sub
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.CreatePackage")
End Sub


Private Function SaveOptionQueries(oPackage As Object, strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    Dim allOK As Boolean
    bResult = True
    allOK = True
    If oPackage.Exists("install") Then
        If oPackage("install").Exists("tables") Then
            Dim oTable As Object
            Dim strOptionQueryFolder As String
            Dim strFilePath As String
            strOptionQueryFolder = strTempFolder & "\lisa\optionqueries"
    
            
            
            For Each oTable In oPackage.Item("install").Item("tables")
                If oTable.Exists("fields") Then
                    Dim oField As Object
                    For Each oField In oTable.Item("fields")
                        
                        If oField.Item("attributes").Item("optionquery") <> "" And oField.Item("attributes").Item("optionquery") <> "''" Then
                            bResult = SaveTextToDisk(oField.Item("attributes").Item("optionquery"), strOptionQueryFolder, oTable.Item("name") & "." & oField.Item("name") & ".txt")
                            
                            If bResult = False Then allOK = False
                            
                        End If
                        'Remove property from JSON object
                        If oField.Item("attributes").Exists("optionquery") Then
                            Call oField.Item("attributes").Remove("optionquery")
                        End If
                    Next
                End If
            Next
        End If
    End If
    SaveOptionQueries = allOK
    
    Exit Function
ErrorHandler:
    Debug.Print Err.Description
    SaveOptionQueries = False
End Function


Private Function SaveSqlOnUpdate(oPackage As Object, strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    Dim allOK As Boolean
    bResult = True
    allOK = True
    If oPackage.Exists("install") Then
        If oPackage("install").Exists("tables") Then
            Dim oTable As Object
            Dim strSqlOnUpdateFolder As String
            Dim strFilePath As String
            strSqlOnUpdateFolder = strTempFolder & "\lisa\sql_on_update"
            
            For Each oTable In oPackage.Item("install").Item("tables")
                If oTable.Exists("fields") Then
                    Dim oField As Object
                    For Each oField In oTable.Item("fields")
                        
                        If oField.Item("attributes").Item("onsqlupdate") <> "" And oField.Item("attributes").Item("onsqlupdate") <> "''" Then
                            bResult = SaveTextToDisk(oField.Item("attributes").Item("onsqlupdate"), strSqlOnUpdateFolder, oTable.Item("name") & "." & oField.Item("name") & ".txt")
                            Call oField.Item("attributes").Remove("onsqlupdate")
                            If bResult = False Then allOK = False
                            
                        End If
                        'Remove property in JSON object
                        If oField.Item("attributes").Exists("onsqlupdate") Then
                            Call oField.Item("attributes").Remove("onsqlupdate")
                        End If
                    Next
                    
                End If
            Next
        End If
    End If
    SaveSqlOnUpdate = allOK
    
    Exit Function
ErrorHandler:
    Debug.Print Err.Description
    SaveSqlOnUpdate = False
End Function


Private Function SaveSqlForNew(oPackage As Object, strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    Dim allOK As Boolean
    bResult = True
    allOK = True
    If oPackage.Exists("install") Then
        If oPackage("install").Exists("tables") Then
            Dim oTable As Object
            Dim strSqlForNewFolder As String
            Dim strFilePath As String
            strSqlForNewFolder = strTempFolder & "\lisa\sql_for_new"
    
            
            
            For Each oTable In oPackage.Item("install").Item("tables")
                If oTable.Exists("fields") Then
                    Dim oField As Object
                    For Each oField In oTable.Item("fields")
                        
                        If oField.Item("attributes").Item("onsqlinsert") <> "" And oField.Item("attributes").Item("onsqlinsert") <> "''" Then
                            bResult = SaveTextToDisk(oField.Item("attributes").Item("onsqlinsert"), strSqlForNewFolder, oTable.Item("name") & "." & oField.Item("name") & ".txt")
                            Call oField.Item("attributes").Remove("onsqlinsert")
                            If bResult = False Then allOK = False
                            
                        End If
                        'Remove property in JSON object
                        If oField.Item("attributes").Exists("onsqlinsert") Then
                            Call oField.Item("attributes").Remove("onsqlinsert")
                        End If
                    Next
                End If
            Next
        End If
    End If
    SaveSqlForNew = allOK
    
    Exit Function
ErrorHandler:
    Debug.Print Err.Description
    SaveSqlForNew = False
End Function


Private Function SaveSqlExpressions(oPackage As Object, strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    Dim allOK As Boolean
    bResult = True
    allOK = True
    If oPackage.Exists("install") Then
        If oPackage("install").Exists("tables") Then
            Dim oTable As Object
            Dim strSqlExpressionsFolder As String
            Dim strFilePath As String
            strSqlExpressionsFolder = strTempFolder & "\lisa\sql_expressions"
    
            
            
            For Each oTable In oPackage.Item("install").Item("tables")
                If oTable.Exists("fields") Then
                    Dim oField As Object
                    For Each oField In oTable.Item("fields")
                        
                        If oField.Item("attributes").Item("sql") <> "" And oField.Item("attributes").Item("sql") <> "''" Then
                            bResult = SaveTextToDisk(oField.Item("attributes").Item("sql"), strSqlExpressionsFolder, oTable.Item("name") & "." & oField.Item("name") & ".txt")
                            Call oField.Item("attributes").Remove("sql")
                            If bResult = False Then allOK = False
                        End If
                        
                        'Remove property in JSON object
                        If oField.Item("attributes").Exists("sql") Then
                            Call oField.Item("attributes").Remove("sql")
                        End If
                    Next
                End If
            Next
        End If
    End If
    SaveSqlExpressions = allOK
    
    Exit Function
ErrorHandler:
    Debug.Print Err.Description
    SaveSqlExpressions = False
End Function


Private Function SaveSqlDescriptive(oPackage As Object, strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    Dim allOK As Boolean
    bResult = True
    allOK = True
    If oPackage.Exists("install") Then
        If oPackage("install").Exists("tables") Then
            Dim oTable As Object
            Dim strSqlDescriptiveFolder As String
            Dim strFilePath As String
            strSqlDescriptiveFolder = strTempFolder & "\lisa\descriptives"
            
            For Each oTable In oPackage.Item("install").Item("tables")
                If oTable.Item("attributes").Item("descriptive") <> "" And oTable.Item("attributes").Item("descriptive") <> "''" Then
                    bResult = SaveTextToDisk(oTable.Item("attributes").Item("descriptive"), strSqlDescriptiveFolder, oTable.Item("name") & ".txt")
                    If bResult = False Then allOK = False
                            
                End If
                
                'Remove property from JSON object
                If oTable.Item("attributes").Exists("descriptive") Then
                    Call oTable.Item("attributes").Remove("descriptive")
                End If
            Next
        End If
    End If
    SaveSqlDescriptive = allOK
    
    Exit Function
ErrorHandler:
    Debug.Print Err.Description
    SaveSqlDescriptive = False
End Function


' ##SUMMARY Saves the specified text in a file in the file system with the specified name in the specified folder.
' Returns true if success and false otherwise.
Private Function SaveTextToDisk(strText As String, strFolderPath As String, strFileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim oStream
    Set oStream = VBA.CreateObject("ADODB.Stream")
    
    Call CreateFolder(strFolderPath, "")
    
    strFileName = strFolderPath & "\" & strFileName
    
    If strText = "" Or strText = "''" Then
        Call Err.Raise(1, , "Empty text was supplied to the stream")
    End If
    
    oStream.Type = adTypeText
    
    oStream.Open
    
    On Error GoTo StreamError
    Call oStream.WriteText(strText)
    Call oStream.SaveToFile(strFileName, adSaveCreateNotExist)
    
    Call oStream.Close
    
    Set oStream = Nothing
    SaveTextToDisk = True
    
    Exit Function
StreamError:
    If Not oStream Is Nothing Then
        If oStream.State = adStateOpen Then oStream.Close
    End If
    
    Set oStream = Nothing
    SaveTextToDisk = False
    
    Exit Function
ErrorHandler:
    Debug.Print "LIPPackageBuilder.SaveTextToDisk " & Err.Description
    SaveTextToDisk = False
End Function


' ##SUMMARY Saves the specified binary data in a file with the specified file name in the specified folder path.
Private Function SaveBinaryToDisk(sBinaryBase64Data As String, sFileName As String, sFolderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Make sure the specified folder exists
    Call CreateFolder(sFolderPath, "")
    
    Dim sFilePath As String
    sFilePath = sFolderPath + "\" + sFileName
    
    Dim binaryData() As Byte
    binaryData = DecodeBase64(sBinaryBase64Data)
    
    Dim binaryStream
    Set binaryStream = VBA.CreateObject("ADODB.Stream")
    binaryStream.Type = adTypeBinary
    
    binaryStream.Open
    
    On Error GoTo StreamError
    binaryStream.Write binaryData
    
    binaryStream.SaveToFile sFilePath, adSaveCreateNotExist
    
    binaryStream.Close
    Set binaryStream = Nothing
    SaveBinaryToDisk = True
    
    Exit Function
StreamError:
    binaryStream.Close
    Set binaryStream = Nothing
    SaveBinaryToDisk = False
    Exit Function
ErrorHandler:
    SaveBinaryToDisk = False
End Function


Private Function SaveTableIcons(oPackage As Object, strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    Dim bAllOK As Boolean
    bResult = True
    bAllOK = True
    
    If oPackage.Exists("install") Then
        If oPackage("install").Exists("tables") Then
            Dim oTable As Object
            Dim strIconFolder As String
            strIconFolder = strTempFolder & "\lisa\icons"
            For Each oTable In oPackage.Item("install").Item("tables")
                If oTable.Exists("attributes") Then
                    If oTable.Item("attributes").Exists("icon") Then
                        bResult = SaveBinaryToDisk(oTable.Item("attributes").Item("icon"), oTable("name") & ".ico", strIconFolder)
                        Call oTable.Item("attributes").Remove("icon")
                        If bResult = False Then bAllOK = False
                    End If
                End If
            Next
        End If
    End If
    SaveTableIcons = bAllOK
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.SaveTableIcons")
End Function


' ##SUMMARY Exports the selected SQL components.
Private Function ExportSql(oPackage As Object, strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    If Not oPackage.Item("install") Is Nothing Then
        Dim oProcedure As Object
        
        If Not oPackage.Item("install").Item("sql") Is Nothing Then
            ' Create folder for sql files
            Dim sSqlFolderPath As String
            sSqlFolderPath = CreateFolder(strTempFolder, "sql")
            For Each oProcedure In oPackage.Item("install").Item("sql")
                bResult = ExportSqlObject(oProcedure.Item("name"), oProcedure.Item("definition"), sSqlFolderPath)
                If bResult = False Then
                    ExportSql = False
                    Exit Function
                End If
                Call oProcedure.Remove("definition")
                Call oProcedure.Add("relPath", "sql\" & oProcedure.Item("name") & ".sql")
            Next
        End If
        
    End If
    
    ExportSql = True
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.ExportSql")
End Function


' ##SUMMARY Exports a single SQL procedure/function as a file in the specified folder.
Private Function ExportSqlObject(ProcedureName As String, Definition As String, sSqlFolderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strFileName As String
    strFileName = sSqlFolderPath & "\" & ProcedureName & ".sql"
    
    Dim strDefinition As String
    strDefinition = StrConv(DecodeBase64(Definition), vbUnicode)
    'Work-around: conversion adds nullchars since it's Unicode (2 bytes), second byte is always null.
    strDefinition = VBA.Replace(strDefinition, Chr(0), "")
    
    Dim intFileNum As Integer
    
    intFileNum = FreeFile
    ' change Output to Append if you want to add to an existing file
    ' rather than creating a new file each time
    Open strFileName For Output As intFileNum
    Print #intFileNum, strDefinition
    Close intFileNum
    
    ExportSqlObject = True

    Exit Function
ErrorHandler:
    ExportSqlObject = False
End Function


Private Function ByteArrayToString(bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    
    ByteArrayToString = sAns
End Function


' ##SUMMARY Lets the user select a folder. Returns the path to the selected folder.
Private Function GetFolder(sText As String) As String
    On Error GoTo ErrorHandler
    
    GetFolder = ""
    
    Dim fDlg As New LCO.FolderDialog
    fDlg.text = sText
    If fDlg.show = VBA.vbOK Then
        GetFolder = fDlg.Folder
    End If
    Set fDlg = Nothing
    
    Exit Function
ErrorHandler:
    GetFolder = ""
    Set fDlg = Nothing
End Function


' ##SUMMARY Zips the content of the specified folder and gives the zip the name specified.
' Returns the full path to the zip file.
Private Function ZipFolder(sZipName As String, sSourcePath As String, sTargetPath As String) As String
    On Error GoTo ErrorHandler
    
    ' Derive the full path for the zip
    Dim sZipFileFullPath As Variant
    sZipFileFullPath = LCO.MakeFileName(sTargetPath, sZipName & ".zip")
    
    ' Make sure there is no preexisting file on that path
    'sZipFileFullPath = CheckUniqueFilename(VBA.CStr(sZipFileFullPath))
    
    'Create empty zip file
    Call NewZip(sZipFileFullPath)
    
    'Create folder object for the zip file
    Dim oApp As Object
    Set oApp = VBA.CreateObject("Shell.Application")
    Close
    Dim oZipFile As Object
    Set oZipFile = oApp.Namespace(sZipFileFullPath)
    
    Dim oPackageFolder As Object
    If Not oZipFile Is Nothing Then
        'Create folder object for the package folder (different path format, which is messed up...)
        Set oPackageFolder = oApp.Namespace(sSourcePath & "\")
        If Not oPackageFolder Is Nothing Then
            'Move files from the package folder to the zip file
            Call oZipFile.CopyHere(oPackageFolder.Items)
        
            'Keep script waiting until Compressing is done
            On Error Resume Next
            Do Until oZipFile.Items.Count = _
               oPackageFolder.Items.Count
                Application.Wait (Now + TimeValue("0:00:01"))
            Loop
            On Error GoTo 0
        Else
            ZipFolder = ""
            Exit Function
        End If
    Else
        ZipFolder = ""
        Exit Function
    End If
    
    ZipFolder = sZipFileFullPath
    
    Exit Function
ErrorHandler:
    ZipFolder = ""
End Function


' ##SUMMARY Checks whether the full file path specified already exists.
' If so, it adds a timestamp stripped of special characters to the end of the file name.
' Returns a guaranteed unique file name, either the original or a fixed version.
'Private Function CheckUniqueFilename(sFullFilePath As String) As String
'    On Error GoTo ErrorHandler
'
'    ' Check if there already is a file on the specified path
'    If VBA.Dir(sFullFilePath) <> "" Then
'        ' Add a clean timestamp to zip file name to make it unique
'        sFullFilePath = VBA.Left(sFullFilePath, VBA.Len(sFullFilePath) - 4) & "_" & GetCleanTimestamp & ".zip"
'    End If
'
'    CheckUniqueFilename = sFullFilePath
'
'    Exit Function
'ErrorHandler:
'    CheckUniqueFilename = ""
'    Call UI.ShowError("LIPPackageBuilder.CheckUniqueFilename")
'End Function


' ##SUMMARY Returns a timestamp where all characters except digits have been removed.
Private Function GetCleanTimestamp() As String
    On Error GoTo ErrorHandler
    
    Dim sResult As String
    sResult = VBA.Now
    
    ' Replace all special characters that are not approved in file names in Windows
    sResult = VBA.Replace(sResult, "\", "")
    sResult = VBA.Replace(sResult, "/", "")
    sResult = VBA.Replace(sResult, ":", "")
    sResult = VBA.Replace(sResult, "*", "")
    sResult = VBA.Replace(sResult, "?", "")
    sResult = VBA.Replace(sResult, """", "")
    sResult = VBA.Replace(sResult, "<", "")
    sResult = VBA.Replace(sResult, ">", "")
    sResult = VBA.Replace(sResult, "|", "")
    
    ' Replace all additional unwanted characters that can be part of a timestamp in different locales
    sResult = VBA.Replace(sResult, "-", "")
    sResult = VBA.Replace(sResult, " ", "")
    sResult = VBA.Replace(sResult, ".", "")
    
    GetCleanTimestamp = sResult

    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.GetCleanTimestamp")
End Function


' ##SUMMARY Renames the folder on the specified path to the name given to the package.
' Returns the full path to the new folder if successful and otherwise an empty string.
Private Function RenameTemporaryFolder(sPackageName As String, sTempFolderPath As String) As String
    On Error GoTo ErrorHandler
    
    ' Remove any trailing \ in the path
    If VBA.Right(sTempFolderPath, 1) = "\" Then
        sTempFolderPath = VBA.Left(sTempFolderPath, VBA.Len(sTempFolderPath) - 1)
    End If
    
    ' Derive the path for the new folder
    Dim sNewFolderPath As String
    sNewFolderPath = VBA.Left(sTempFolderPath, VBA.InStrRev(sTempFolderPath, "\")) & sPackageName
    
    ' If there already is a folder on the desired path, delete it (safe since this is within the designated folder for LIPPackageBuilder within the Lime CRM temporary folder.
    If VBA.Dir(sNewFolderPath, vbDirectory) <> "" Then
        Call DeleteTemporaryFolder(sNewFolderPath)
    End If
    
    ' Rename the folder to the package name
    Name sTempFolderPath As sNewFolderPath
    
    RenameTemporaryFolder = sNewFolderPath
    
    Exit Function
ErrorHandler:
    RenameTemporaryFolder = ""
End Function


Private Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    Dim fNum As Integer
    fNum = FreeFile
    
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #fNum
    Print #fNum, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #fNum
End Sub


Private Function DeleteTemporaryFolder(strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler

    'Delete all files and subfolders
    'Be sure that no file is open in the folder
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Right(strTempFolder, 1) = "\" Then
        strTempFolder = Left(strTempFolder, Len(strTempFolder) - 1)
    End If

    If fso.FolderExists(strTempFolder) = False Then
        DeleteTemporaryFolder = True
        Exit Function
    End If

    On Error Resume Next
    'Delete files
    fso.DeleteFile strTempFolder & "\*.*", True
    'Delete subfolders
    fso.DeleteFolder strTempFolder & "\*.", True
    Call RmDir(strTempFolder)
    On Error GoTo 0
    
    DeleteTemporaryFolder = True
    
    Exit Function
ErrorHandler:
    DeleteTemporaryFolder = False
    Debug.Print Err.Number & vbCrLf & Err.Description
End Function


Private Function SavePackageFile(oPackage As Object, strTempPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    Dim fso As New FileSystemObject
    Dim filePath As String
    filePath = strTempPath & "\package.json"
    bResult = True
    'Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(filePath, True, False)
    'Convert to a string and save
    Call oFile.WriteLine(JsonConverter.ConvertToJson(oPackage))
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
    
    SavePackageFile = bResult
    
    Exit Function
ErrorHandler:
    bResult = False
End Function


' ##SUMMARY Exports all Actionpads included in the Package JSON.
Private Function ExportActionpads(ByRef oPackage As Object, sTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not oPackage.Item("install") Is Nothing Then
        If Not oPackage.Item("install").Item("actionpads") Is Nothing Then
            Dim oActionpad As Object
            Dim sActionpadsFolderPath As String
            sActionpadsFolderPath = CreateFolder(sTempFolder, "actionpads")
            For Each oActionpad In oPackage.Item("install").Item("actionpads")
                Call VBA.FileCopy(LCO.MakeFileName(Application.WebFolder, oActionpad.Item("fileName")), LCO.MakeFileName(sActionpadsFolderPath, oActionpad.Item("fileName")))
            Next
        End If
    End If
    ExportActionpads = True
    
    Exit Function
ErrorHandler:
    ExportActionpads = False
End Function


' ##SUMMARY Exports all VBA modules marked in the Package JSON.
Private Function ExportVBA(oPackage As Object, strTempFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    bResult = True
    If Not oPackage.Item("install") Is Nothing Then
        Dim oModule As Object
        
        If Not oPackage.Item("install").Item("vba") Is Nothing Then
            Dim sVBAFolderPath As String
            sVBAFolderPath = CreateFolder(strTempFolder, "vba")
            For Each oModule In oPackage.Item("install").Item("vba")
                
                bResult = ExportVBAModule(oModule.Item("name"), sVBAFolderPath)
                If bResult = False Then
                    ExportVBA = False
                    Exit Function
                End If
            Next
        End If
    End If
    ExportVBA = bResult
    
    Exit Function
ErrorHandler:
    ExportVBA = False
End Function


' ##SUMMARY Exports a VBA module to a file.
Private Function ExportVBAModule(ModuleName As String, sVBAFolderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bResult As Boolean
    bResult = True
    
    ' Retrieve the VBA code file
    Dim strFileName As String
    Dim Component As Object
    Set Component = ThisApplication.VBE.ActiveVBProject.VBComponents(ModuleName)
    If Not Component Is Nothing Then
        strFileName = Component.Name
        Select Case Component.Type
            Case 1
                strFileName = strFileName & ".bas"
            Case 2
                strFileName = strFileName & ".cls"
            Case 3
                strFileName = strFileName & ".frm"
            
            Case Else
                bResult = False
                Exit Function
        End Select
        
        Call Component.Export(sVBAFolderPath & "\" & strFileName)
        bResult = True
    End If
    
    ExportVBAModule = bResult
    
    Exit Function
ErrorHandler:
    ExportVBAModule = False
End Function


' ##SUMMARY Checks if the specified sub folder exists in the specified parent folder and creates it otherwise.
' If the parent folder does not exist the function calls itself recursively to create the parent folder first.
' Can also be called with an empty string as subfolder. Will then make sure that the parent folder exists and create it otherwise.
' Returns the full path for the parent folder plus sub folder.
Private Function CreateFolder(sParentFolderPath As String, sSubFolderName As String) As String
    On Error GoTo ErrorHandler
    
    Dim sFullPath As String
    
    ' Fix parent folder path
    If VBA.Right(sParentFolderPath, 1) = "\" Then
        sParentFolderPath = VBA.Left(sParentFolderPath, VBA.Len(sParentFolderPath) - 1)
    End If
    
    ' Check if parent folder exists
    If VBA.Dir(sParentFolderPath, VBA.vbDirectory) = "" Then
        ' Create parent folder recursively
        Call CreateFolder(VBA.Left(sParentFolderPath, VBA.InStrRev(sParentFolderPath, "\") - 1), _
                VBA.Right(sParentFolderPath, VBA.Len(sParentFolderPath) - VBA.InStrRev(sParentFolderPath, "\")))
    End If
    
    If sSubFolderName <> "" Then
        ' Check if sub folder exists and create otherwise
        sFullPath = VBA.Replace(sParentFolderPath & "\" & sSubFolderName, "\\", "\")
        If VBA.Dir(sFullPath, VBA.vbDirectory) = "" Then
            Call VBA.MkDir(sFullPath)
        End If
    Else
        sFullPath = sParentFolderPath
    End If
    
    CreateFolder = sFullPath
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.CreateFolder")
End Function


' ##SUMMARY Called from javascript.
Public Function GetLocalizations(ByVal sOwner As String) As Records
    On Error GoTo ErrorHandler
    Dim oRecords As New LDE.Records
    Dim oFilter As New LDE.Filter
    Dim oView As New LDE.View
    
    Call oView.Add("owner", lkSortAscending)
    Call oView.Add("code")
    Call oView.Add("context")
    Call oView.Add("sv")
    Call oView.Add("en_us")
    Call oView.Add("no")
    Call oView.Add("fi")
    Call oView.Add("da")
    If sOwner <> "" Then
        Call oFilter.AddCondition("owner", lkOpEqual, sOwner)
    End If
    Call oRecords.Open(Database.Classes("localize"), oFilter, oView)
    Set GetLocalizations = oRecords
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.GetLocalizations")
End Function


' ##SUMMARY Called from javascript.
Public Function OpenExistingPackage() As String
    On Error GoTo ErrorHandler
    
    Dim o As New LCO.FileOpenDialog
    Dim strFilePath As String
    o.AllowMultiSelect = False
    o.Caption = "Select Package file"
    o.Filter = "Zipped Package files (*.zip) | *.zip"

    o.DefaultFolder = LCO.GetDesktopPath
    If o.show = vbOK Then
        strFilePath = o.FileName
    Else
        Exit Function
    End If

    If LCO.ExtractFileExtension(strFilePath) = "zip" Then
        Dim strTempFolderPath As String
        strTempFolderPath = Application.TemporaryFolder & "\" & VBA.Replace(VBA.Replace(LCO.GenerateGUID, "{", ""), "}", "")
        Dim fso As New Scripting.FileSystemObject
        If Not fso.FolderExists(strTempFolderPath) Then
            Call fso.CreateFolder(strTempFolderPath)
        End If


        On Error GoTo UnzipError
        Call UnZip(strTempFolderPath, strFilePath)

        On Error GoTo ErrorHandler
        Dim strJson As String
        If LCO.FileExists(strTempFolderPath & "\" & "app.json") Then
            strJson = ReadAllTextFromFile(strTempFolderPath & "\" & "app.json")
        ElseIf LCO.FileExists(strTempFolderPath & "\" & "package.json") Then
            strJson = ReadAllTextFromFile(strTempFolderPath & "\" & "package.json")
        Else
            Call Application.MessageBox("Could not find an app.json or a package.json in the extracted folder")
            Exit Function
        End If

        Dim b64Json As String
        b64Json = XMLEncodeBase64(strJson)

        OpenExistingPackage = b64Json
    End If
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.OpenExistingPackage")
    Exit Function
UnzipError:
    Call Application.MessageBox("There was an error unzipping the zipped package file")
End Function


' ##SUMMARY Called from javascript.
' Lets the user select an existing CHANGELOG.md file. Then it retrieves information about
' the version number and authors of the latest version.
' If the user cancels the file dialog then an empty JSON object will be returned.
Public Function OpenExistingChangelog() As String
    On Error GoTo ErrorHandler

    ' Let the user select a file
    '##TODO
    
    ' Read line by line until the wanted information has been found.
    '##TODO
    
    ' Build JSON containing the desired information
    Dim sChangelogInfoJson As String
    sChangelogInfoJson = "{" _
                            & """versionNumber"" : """ & "1.2.3" & """," _
                            & """authors"" : """ & "FER" & """" _
                        & "}"
    
    OpenExistingChangelog = sChangelogInfoJson

    Exit Function
ErrorHandler:
    OpenExistingChangelog = "{}"
    Call UI.ShowError("LIPPackageBuilder.OpenExistingChangelog")
End Function


Private Sub UnZip(strTargetPath As String, Fname As Variant)
    Dim oApp As Object, FSOobj As Object
    Dim FileNameFolder As Variant

    If Right(strTargetPath, 1) <> "\" Then
        strTargetPath = strTargetPath & "\"
    End If
    
    FileNameFolder = strTargetPath
    
    'create destination folder if it does not exist
    Set FSOobj = CreateObject("Scripting.FilesystemObject")
    If FSOobj.FolderExists(FileNameFolder) = False Then
        FSOobj.CreateFolder FileNameFolder
    End If
    
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(CVar(Fname)).Items
    
    Set oApp = Nothing
    Set FSOobj = Nothing
    Set FileNameFolder = Nothing
    
End Sub


Private Function ReadAllTextFromFile(strFilePath As String) As String
On Error GoTo ErrorHandler
    Dim strText As String, Filenum As Integer, s As String
    
    Filenum = FreeFile
    
    Open strFilePath For Input As #Filenum

    While Not EOF(Filenum)
        Line Input #Filenum, s    ' read in data 1 line at a time

        strText = strText + s + VBA.vbNewLine
    Wend
    
    Close #Filenum
    
    ReadAllTextFromFile = strText
Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.ReadAllTextFromFile")
End Function


Private Function CleanupPackageFile(oPackage As Object) As Boolean
On Error GoTo ErrorHandler
    'Remove related table
     If oPackage.Exists("install") Then
        If oPackage("install").Exists("tables") Then
            Dim oTable As Object
            Dim oField As Object
            
            For Each oTable In oPackage.Item("install").Item("tables")
                If oTable.Exists("fields") Then
                    For Each oField In oTable.Item("fields")
                        If oField.Exists("attributes") Then
                            Call oField.Item("attributes").Remove("relatedtable")
                        End If
                    Next
                End If
            Next
        End If
    End If
    CleanupPackageFile = True
Exit Function
ErrorHandler:
    CleanupPackageFile = False
End Function


' ##SUMMARY Transforms the input string to a JSON object and returns that object.
' The input parameter must be a base64 encoded string.
Private Function Base64StringToJsonObject(sInput As String) As Object
    On Error GoTo ErrorHandler

    Dim sJson As String
    sJson = VBA.StrConv(DecodeBase64(sInput), VBA.vbUnicode)
    Set Base64StringToJsonObject = JsonConverter.ParseJson(sJson)
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.Base64StringToJsonObject")
End Function


' ##SUMMARY Copies all the content in the specified source path to the specified target path.
' Returns true if successful and false if any error was encountered.
Private Function CopyFolder(sSourcePath As String, sTargetPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim objFSO As Object
    Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")
    Call objFSO.CopyFolder(sSourcePath, sTargetPath)       'object.CopyFolder SOURCE, DESTINATION[, OVERWRITE]. Default of OVERWRITE is true.

    CopyFolder = True

    Exit Function
ErrorHandler:
    CopyFolder = False
    Call UI.ShowError("LIPPackageBuilder.CopyFolder")
End Function


' ##SUMMARY Creates the README.md file needed for add-ons. Will replace placeholders with data from the specified oReadmeInfo JSON.
Private Function CreateReadmeMd(ByRef oReadmeInfo As Object, sPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim sReadme As String
    sReadme = ReadAllTextFromFile(Application.WebFolder & "apps\LIPPackageBuilder\templates\README.md")
    
    sReadme = VBA.Replace(sReadme, "<*displayName*>", oReadmeInfo.Item("displayName"))
    sReadme = VBA.Replace(sReadme, "<*description*>", oReadmeInfo.Item("description"))
    If Not oReadmeInfo.Item("cloudCompatible") Then
        sReadme = VBA.Replace(sReadme, "<*cloudCompatible*>", "* This add-on is not compatible with the Lime CRM Cloud environment.")
    Else
        sReadme = VBA.Replace(sReadme, "<*cloudCompatible*>", "")
    End If
    Call SaveTextToDisk(sReadme, sPath, "README.md")
    
    CreateReadmeMd = True

    Exit Function
ErrorHandler:
    CreateReadmeMd = False
    Call UI.ShowError("LIPPackageBuilder.CreateReadmeMd")
End Function


' ##SUMMARY Creates the README.md file needed for add-ons. Will replace placeholders with data from the specified oReadmeInfo JSON.
Private Function CreateChangelogMd(ByRef oChangelogInfo As Object, sPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim sChangelog As String
    sChangelog = ReadAllTextFromFile(Application.WebFolder & "apps\LIPPackageBuilder\templates\CHANGELOG.md")
    
    sChangelog = VBA.Replace(sChangelog, "<*displayName*>", oChangelogInfo.Item("displayName"))
    sChangelog = VBA.Replace(sChangelog, "<*versionNumber*>", oChangelogInfo.Item("versionNumber"))
    sChangelog = VBA.Replace(sChangelog, "<*date*>", oChangelogInfo.Item("date"))
    sChangelog = VBA.Replace(sChangelog, "<*authors*>", oChangelogInfo.Item("authors"))
    
    ' Add an * at the beginning of the version comments if not already there.
    Dim comments As String
    comments = oChangelogInfo.Item("versionComment")
    If comments <> "" Then
        If Not VBA.Left(comments, 1) = "*" Then
            comments = "*" & comments
        End If
    End If
    sChangelog = VBA.Replace(sChangelog, "<*versionComment*>", comments)
    
    Call SaveTextToDisk(sChangelog, sPath, "CHANGELOG.md")
    
    CreateChangelogMd = True

    Exit Function
ErrorHandler:
    CreateChangelogMd = False
    Call UI.ShowError("LIPPackageBuilder.CreateChangelogMd")
End Function


' ##SUMMARY Called from javascript. Returns the version of the LIP Package Builder as a string.
Public Function GetVersion() As String
    On Error GoTo ErrorHandler

    GetVersion = m_sLIPPackageBuilderVersion

    Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.GetVersion")
End Function

