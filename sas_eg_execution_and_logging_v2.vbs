'00_Setup
Option Explicit
On Error Resume Next

Dim application
Dim project_name
Dim project

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

project_name = WScript.Arguments.Item(0)

'Confirm the existence of the project
If WScript.Arguments.Count = 0 Then 
    WScript.Echo "ERROR: Expecting the full path name of a project file"
    WScript.Quit -1
End If


'01_Start up SAS EG application
Set application = CreateObject("SASEGObjectModel.Application.8.1")
If fn_Check_Error("CreateObject") = True Then
    WScript.Quit -1
End If


'02_Open your SAS EG project with the application
Set project = application.Open(project_name,"")
If fn_Check_Error("application.Open") = True Then
    WScript.Quit -1
End If
WScript.Echo "Opening: " & project_name

    
'03_Run the project
project.run
If fn_Check_Error("project.run") = True Then
    WScript.Quit -1
End If
WScript.Echo "Execution Completes."


'04_Save the project
project.Save
If fn_Check_Error("project.Save") = True Then
    WScript.Quit -1
End If
WScript.Echo "Save successfully."  


'05_Save all the available logs and check if any error occurs in the project by node
Dim folder_name_project 
Dim table_of_contents
Dim file_name
Dim flow
Dim item

folder_name_project = fn_Get_Working_Directory() & "\SAS_EG_Logs_" & project.Name
fn_Create_Folder(folder_name_project)
Wscript.Echo "Logs are saved at " & folder_name_project

table_of_contents = "Table of Contents" & vbNewLine & "-" & vbNewLine

'Loop through all the flows and nodes in the current project
For Each flow In project.ContainerCollection
    If flow.ContainerType = 0 Then 'Process flow is containerType of 0
        Dim folder_name_flow
        Dim count
        
        folder_name_flow = folder_name_project & "\" & flow.Name
        fn_Create_Folder(folder_name_flow)

        table_of_contents = "Flow Name: " & flow.Name & vbNewLine & "---" & vbNewLine
        count = 0
        
        For Each item in flow.Items
            count = count + 1
            table_of_contents = table_of_contents & count & "_" & item.Name & vbNewLine

            Select Case item.Type
                Case egQuery
                    'Do nothing。這邊以eqQuery為例說明，你可以用這行程式碼:Wscript.Echo item.Type，來判斷你的EG item
                Case Else
                    If (Not item.Log Is Nothing) Then
                        file_name = folder_name_flow & "\" & count & "_" & item.Name & ".log"
                        fn_Save_As_Log item.Log.Text, file_name
                        fn_Check_String_Error(file_name)
                    End If
            End Select
        Next            
        
        fn_Save_As_Log table_of_contents, folder_name_flow & "\" & flow.Name & ".txt" 
    End If
Next


'06_Shut down the application
If not (application Is Nothing) Then
    application.Quit
    Set application = Nothing
End If


'------------------Helper functions------------------
Function fn_Check_Error(functionName)   
    Dim strmsg, errNum
    fn_Check_Error = False

    If Err.Number <> 0 Then
        strmsg = "Error #" & Hex(Err.Number) & vbCrLf & "In Function " & functionName & vbCrLf & Err.Description
        MsgBox strmsg  'Get notified via MessageBox of Errors in the script.
        fn_Check_Error = True
    End If
End Function

'Function to get the current working directory
Function fn_Get_Working_Directory()
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    fn_Get_Working_Directory = objShell.CurrentDirectory
End Function

'Function to create a new folder
Function fn_Create_Folder(folderName)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(folderName) = False Then
        objFSO.CreateFolder folderName
    End If
End Function

'Function to save the text as a .log file
Function fn_Save_As_Log(text, fileName)
    Dim objFS
    Dim objOutFile
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objOutFile = objFS.CreateTextFile(fileName, True)

    objOutFile.Write(text)
    objOutFile.Close
End Function

'Function to check if there is a "ERROR" string in the log
Function fn_Check_String_Error(fileName)
    Dim objFSO, objInputFile, readAll
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objInputFile = objFSO.OpenTextFile(fileName)
    readAll = objInputFile.ReadAll

    If InStr(readAll, "ERROR") <> 0 Then
        'Formulate your next thing here if any error occurs.
        WScript.Echo "There are errors in " & fileName
    Else
        WScript.Echo "There is no error in " & fileName
    End If
End Function