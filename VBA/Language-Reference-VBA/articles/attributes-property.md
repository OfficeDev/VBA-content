---
title: Attributes Property
keywords: vblr6.chm2181972
f1_keywords:
- vblr6.chm2181972
ms.prod: office
api_name:
- Office.Attributes
ms.assetid: 965b8844-33b8-2358-5bdd-cc814987132f
ms.date: 06/08/2017
---


# Attributes Property



 **Description**
Sets or returns the attributes of files or folders. Read/write or read-only, depending on the attribute.
 **Syntax**
 _object_. **Attributes** [= _newattributes_ ]
The  **Attributes** property has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **File** or **Folder** object.|
| _newattributes_|Optional. If provided,  _newattributes_ is the new value for the attributes of the specified _object_.|
 **Settings**
The  _newattributes_ argument can have any of the following values or any logical combination of the following values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**Normal**|0|Normal file. No attributes are set.|
|**ReadOnly**|1|Read-only file. Attribute is read/write.|
|**Hidden**|2|Hidden file. Attribute is read/write.|
|**System**|4|System file. Attribute is read/write.|
|**Volume**|8|Disk drive volume label. Attribute is read-only.|
|**Directory**|16|Folder or directory. Attribute is read-only.|
|**Archive**|32|File has changed since last backup. Attribute is read/write.|
|**Alias**|64|Link or shortcut. Attribute is read-only.|
|**Compressed**|128|Compressed file. Attribute is read-only.|
 **Remarks**
The following code illustrates the use of the  **Attributes** property with a file:



```vb
Sub SetClearArchiveBit(filespec)
    Dim fs, f, r
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(fs.GetFileName(filespec))
    If f.attributes and 32 Then
        r = MsgBox("The Archive bit is set, do you want to clear it?", vbYesNo, "Set/Clear Archive Bit")
        If r = vbYes Then 
            f.attributes = f.attributes - 32
            MsgBox "Archive bit is cleared."
        Else
            MsgBox "Archive bit remains set."
        End If
    Else
        r = MsgBox("The Archive bit is not set. Do you want to set it?", vbYesNo, "Set/Clear Archive Bit")
        If r = vbYes Then 
f.attributes = f.attributes + 32
            MsgBox "Archive bit is set."
        Else
            MsgBox "Archive bit remains clear."
        End If
    End If
End Sub
```


