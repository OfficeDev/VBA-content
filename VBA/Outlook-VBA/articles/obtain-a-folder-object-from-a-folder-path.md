---
title: Obtain a Folder Object from a Folder Path
ms.prod: outlook
ms.assetid: c576924a-6bf9-7bae-bcee-7bacd299e144
ms.date: 06/08/2017
---


# Obtain a Folder Object from a Folder Path

This topic shows a function that accepts a folder path and returns a  **[Folder](folder-object-outlook.md)** object that corresponds to the specified folder. For example, if you provide the folder path "Mailbox - Dan Wilson\Inbox\Customers", the code in the `TestGetFolder` procedure will display the **Folder** object that corresponds to the Customers folder under Dan Wilson's Inbox, if the Customers folder exists under the Inbox. If the Customers folder does not exist, `GetFolder` will return `Nothing`.


```vb
Function GetFolder(ByVal FolderPath As String) As Outlook.Folder 
 Dim TestFolder As Outlook.Folder 
 Dim FoldersArray As Variant 
 Dim i As Integer 
 
 On Error GoTo GetFolder_Error 
 If Left(FolderPath, 2) = "\\" Then 
 FolderPath = Right(FolderPath, Len(FolderPath) - 2) 
 End If 
 'Convert folderpath to array 
 FoldersArray = Split(FolderPath, "\") 
 Set TestFolder = Application.Session.Folders.item(FoldersArray(0)) 
 If Not TestFolder Is Nothing Then 
 For i = 1 To UBound(FoldersArray, 1) 
 Dim SubFolders As Outlook.Folders 
 Set SubFolders = TestFolder.Folders 
 Set TestFolder = SubFolders.item(FoldersArray(i)) 
 If TestFolder Is Nothing Then 
 Set GetFolder = Nothing 
 End If 
 Next 
 End If 
 'Return the TestFolder 
 Set GetFolder = TestFolder 
 Exit Function 
 
GetFolder_Error: 
 Set GetFolder = Nothing 
 Exit Function 
End Function 
 
Sub TestGetFolder() 
 Dim folder As Outlook.Folder 
 Set folder = GetFolder ("\\Mailbox - Dan Wilson\Inbox\Customers") 
 If Not(folder Is Nothing) Then 
 folder.Display 
 End If 
End Sub
```


