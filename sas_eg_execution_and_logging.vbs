Option Explicit
On Error Resume Next

Dim appObject
Dim projObject
Dim prjName 

prjName = WScript.Arguments.Item(0)

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "ERROR: Expecting the full path name of a project file"
    WScript.Quit -1
End If

Call execute_and_save_project()

Call retrive_project_log()

Sub execute_and_save_project()
    Set appObject = CreateObject("SASEGObjectModel.Application.8.1")
    If fn_Check_Error("CreateObject") = True Then
        Exit Sub
    End If

    Set projObject = appObject.Open(prjName,"")
    If fn_Check_Error("appObject.Open") = True Then
        Exit Sub
    End If

    projObject.run
    If fn_Check_Error("Project.run") = True Then
        Exit Sub
    End If

    projObject.Save
    If fn_Check_Error("Project.Save") = True Then
        Exit Sub
    End If

    projObject.Close
    If fn_Check_Error("Project.Close") = True Then
        Exit Sub
    End If

    If not (appObject Is Nothing) Then
        appObject.Quit
        Set appObject = Nothing
    End If
End Sub




sub retrive_project_log()
    Set appObject = CreateObject("SASEGObjectModel.Application.8.1")
    If fn_Check_Error("CreateObject") = True Then
        Exit Sub
    End If

    Set projObject = appObject.Open(prjName,"")
    If fn_Check_Error("appObject.Open") = True Then
        Exit Sub
    End If

    Const egLog = 0  
    Const egCode = 1  
    Const egData = 2  
    Const egQuery = 3  
    Const egContainer = 4  
    Const egDocBuilder = 5  
    Const egNote = 6  
    Const egResult = 7  
    Const egTask = 8  
    Const egTaskCode = 9  
    Const egProjectParameter = 10  
    Const egOutputData = 11  
    Const egStoredProcess = 12  
    Const egStoredProcessParameter = 13  
    Const egPublishAction = 14  
    Const egCube = 15  
    Const egReport = 18  
    Const egReportSnapshot = 19  
    Const egOrderedList = 20  
    Const egSchedule = 21  
    Const egLink = 22  
    Const egFile = 23  
    Const egIntrNetApp = 24  
    Const egInformationMap = 25  

    Dim projFolder
    Dim flowFolder
    Dim flow
    Dim node
    Dim tableOfContents
    Dim count

    projFolder = fn_Get_Working_Directory() & "\SAS_EG_Logs_" & projObject.Name
    fn_Create_Folder(projFolder)
    Wscript.Echo "Logs are saved at " & projFolder

    For Each flow In projObject.ContainerCollection
        If flow.ContainerType = 0 Then


            flowFolder = projFolder & "\" & flow.Name	
            fn_Create_Folder(flowFolder)

            count = 0


            tableOfContents = "Flow Name: " & flow.Name & vbNewLine & "---" & vbNewLine

            For Each node in flow.Items
                
                count = count + 1

                Select Case node.Type
                    Case egFile 
                        tableOfContents = tableOfContents & "[File]  " & count & "_" & node.Name & vbNewLine

                    Case egCode 
                        tableOfContents = tableOfContents & "[Code]  " & count & "_" & node.Name & vbNewLine

                        If (Not node.Log Is Nothing) Then
                            call fn_Save_As_Log(flowFolder & "\" & count & "_" & node.Name & ".log", node.Log.Text)
                        End If

                    Case egData
                        tableOfContents = tableOfContents & "[Data]  " & count & "_" & node.Name & vbNewLine

                    Case egTask
                        tableOfContents = tableOfContents & "[Task]  " & count & "_" & node.Name & vbNewLine	

                        If (Not node.Log Is Nothing) Then
                            call fn_Save_As_Log(flowFolder & "\" & count & "_" & node.Name & ".log", node.Log.Text)
                        End If

                    Case egQuery
                        tableOfContents = tableOfContents & "[Query] " & count & "_" & node.Name & vbNewLine

                        If (Not node.Log Is Nothing) Then
                            call fn_Save_As_Log(flowFolder & "\" & count & "_" & node.Name & ".log", node.Log.Text)
                        End If
                End Select
            Next

            fn_Save_As_Log flowFolder & "\" & flow.Name & ".txt", tableOfContents
        End If
    Next
    
    projObject.Close
    If fn_Check_Error("Project.Close") = True Then
        Exit Sub
    End If

    If not (appObject Is Nothing) Then
        appObject.Quit
        Set appObject = Nothing
    End If
End Sub

'---------------- Helper functions ----------------
'Function to check if any error occurs in the current function
Function fn_Check_Error(functionName)   
    Dim strmsg, errNum
    fn_Check_Error = False

    If Err.Number <> 0 Then
        strmsg = "Error #" & Hex(Err.Number) & vbCrLf & "In Function " & functionName & vbCrLf & Err.Description
        MsgBox strmsg  'Get notified via message box of errors in the script.
        fn_Check_Error = True
    End If
End Function

'Function to get the current working directory
Function fn_Get_Working_Directory()
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    fn_Get_Working_Directory = objShell.CurrentDirectory
End Function

'Function to save the text as a .log file
Function fn_Save_As_Log(fileName, text)
    Dim objFS
    Dim objOutFile
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objOutFile = objFS.CreateTextFile(fileName, True)

    objOutFile.Write(text)
    objOutFile.Close
End Function

'Function to create a new folder
Function fn_Create_Folder(folderName)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(folderName) = False Then
        objFSO.CreateFolder folderName
    End If
End Function

