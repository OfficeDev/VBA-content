---
title: FileDialogFilter.Description Property (Office)
keywords: vbaof11.chm254003
f1_keywords:
- vbaof11.chm254003
ms.prod: office
api_name:
- Office.FileDialogFilter.Description
ms.assetid: ae3c17d7-62e7-21f5-b543-ee498b7f4d23
ms.date: 06/08/2017
---


# FileDialogFilter.Description Property (Office)

Gets the description of each  **Filter** object as a **String** value. The description is the text that is displayed in a file dialog box. Read-only.


## Syntax

 _expression_. **Description**

 _expression_ Required. A variable that represents a **[FileDialogFilter](filedialogfilter-object-office.md)** object.


## Example

The following example iterates through the default filters of the  **SaveAs** dialog box and displays the description of each filter that includes a Microsoft Excel file. The Extensions property is used to find the appropriate filter objects.


```
Sub Main() 
 
 'Declare a variable as a FileDialogFilters collection. 
 Dim fdfs As FileDialogFilters 
 
 'Declare a variable as a FileDialogFilter object. 
 Dim fdf As FileDialogFilter 
 
 'Set the FileDialogFilters collection variable to 
 'the FileDialogFilters collection of the SaveAs dialog box. 
 Set fdfs = Application.FileDialog(msoFileDialogSaveAs).Filters 
 
 'Iterate through the description and extensions of each 
 'default filter in the SaveAs dialog box. 
 For Each fdf In fdfs 
 
 'Display the description of filters that include 
 'Microsoft Excel files. 
 If InStr(1, fdf.Extensions, "xls", vbTextCompare) > 0 Then 
 MsgBox "Filter description: " &amp; fdf.Description 
 End If 
 Next fdf 
 
End Sub
```


## See also


#### Concepts


[FileDialogFilter Object](filedialogfilter-object-office.md)
#### Other resources


[FileDialogFilter Object Members](filedialogfilter-members-office.md)

