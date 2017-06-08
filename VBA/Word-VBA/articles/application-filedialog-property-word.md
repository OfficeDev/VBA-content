---
title: Application.FileDialog Property (Word)
keywords: vbawd10.chm158335426
f1_keywords:
- vbawd10.chm158335426
ms.prod: word
api_name:
- Word.Application.FileDialog
ms.assetid: ef478a81-db1d-4bf4-a146-3ff7dd84116b
ms.date: 06/08/2017
---


# Application.FileDialog Property (Word)

Returns a  **FileDialog** object which represents a single instance of a file dialog box.


## Syntax

 _expression_ . **FileDialog**( **_FileDialogType_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileDialogType_|Required| **MsoFileDialogType**|The type of dialog.|

## Example

This example displays the  **Save As** dialog box.


```vb
Sub ShowSaveAsDialog() 
 Dim dlgSaveAs As FileDialog 
 Set dlgSaveAs = Application.FileDialog( _ 
 FileDialogType:=msoFileDialogSaveAs) 
 dlgSaveAs.Show 
End Sub
```

This example displays the  **Open** dialog box and allows a user to select multiple files to open.




```vb
Sub ShowFileDialog() 
 Dim dlgOpen As FileDialog 
 Set dlgOpen = Application.FileDialog( _ 
 FileDialogType:=msoFileDialogOpen) 
 With dlgOpen 
 .AllowMultiSelect = True 
 .Show 
 End With 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

