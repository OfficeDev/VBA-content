---
title: Name Property (FileSystemObject object)
keywords: vblr6.chm2181996
f1_keywords:
- vblr6.chm2181996
ms.prod: office
ms.assetid: 1e2c7813-74da-fd24-4e2f-4855f2d57015
ms.date: 06/08/2017
---


# Name Property (FileSystemObject object)



 **Description**
Sets or returns the name of a specified file or folder. Read/write.
 **Syntax**
 _object_. **Name** [= _newname_ ]
The  **Name** property has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **File** or **Folder** object.|
| _newname_|Optional. If provided,  _newname_ is the new name of the specified _object_.|
 **Remarks**
The following code illustrates the use of the  **Name** property:



```vb
Sub ShowFileAccessInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.Name &; " on Drive " &; UCase(f.Drive) &; vbCrLf
    s = s &; "Created: " &; f.DateCreated &; vbCrLf
    s = s &; "Last Accessed: " &; f.DateLastAccessed &; vbCrLf
    s = s &; "Last Modified: " &; f.DateLastModified  
    MsgBox s, 0, "File Access Info"
End Sub
```


