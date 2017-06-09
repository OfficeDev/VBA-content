---
title: Request access to multiple files
ms.prod: office
ms.date: 06/08/2017
---
# Request access to multiple files

Use the **GrantAccessToMultipleFiles** command to request access to multiple files at once in your Office 2016 for Mac solution.

***Applies to:*** *Excel for Mac | PowerPoint for Mac | Word for Mac | Office 2016 for Mac*

Unlike VBA macros in Office for Mac 2011, VBA macros in Office 2016 for Mac don’t have access to external files by default. Because the Office 2016 for Mac apps are sandboxed, they do not have permission to access external files. 

Existing macro file commands prompt the user for permission to access a file if the app doesn’t have access to it. This means that macros that access external files cannot run unattended. The user must approve file access the first time each file is referenced. You can use the **GrantAccessToMultipleFiles** command to minimize the number of prompts in order to improve the user experience. 

## GrantAccessToMultipleFiles command
Use the **GrantAccessToMultipleFiles** command to input an array of file paths and prompt the user for permission to access them.

```vb
    Boolean  GrantAccessToMultipleFiles(fileArray) 
```

|**Parameter**|**Description**|
|:-----|:-----|
|*fileArray*|An array of POSIX file paths|

The command returns whether the user granted permission or not.

|**Return value**|**Description**|
|:-----|:-----|
|True|The user grants permission to the files.|
|False|The user denies permission to the files.|

**Note:** After the user grants permissions, the permissions are stored with the app. The user doesn’t need to grant permission to the file again. 

**Example**

```vb
    Sub requestFileAccess()  

    'Declare Variables  
    Dim fileAccessGranted As Boolean  
    Dim filePermissionCandidates 
  
   'Create an array with file paths for the permissions that are needed.  
    filePermissionCandidates = Array("/Users//Desktop/test1.txt", "/Users//Desktop/test2.txt") 
  
    'Request access from user.  
     fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates) 
    'Returns true if access is granted; otherwise, false. 
    End Sub
```
