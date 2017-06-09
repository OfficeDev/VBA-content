---
title: Application.FileDialog Property (Excel)
keywords: vbaxl10.chm133270
f1_keywords:
- vbaxl10.chm133270
ms.prod: excel
api_name:
- Excel.Application.FileDialog
ms.assetid: 96a6fdc5-1bde-68dd-2493-9d8a92915afb
ms.date: 06/08/2017
---


# Application.FileDialog Property (Excel)

Returns a  **[FileDialog](http://msdn.microsoft.com/library/71a030f2-3b02-21e1-c156-0514ff5eddb7%28Office.15%29.aspx)** object representing an instance of the file dialog.


## Syntax

 _expression_ . **FileDialog**( **_fileDialogType_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fileDialogType_|Required| **[MsoFileDialogType](http://msdn.microsoft.com/library/ee445a67-1193-f446-4bd2-963c07fba5ae%28Office.15%29.aspx)**|The type of file dialog.|

## Remarks





| **MsoFileDialogType** can be one of these **MsoFileDialogType** constants.|
| **msoFileDialogFilePicker** . Allows user to select a file.|
| **msoFileDialogFolderPicker** . Allows user to select a folder.|
| **msoFileDialogOpen** . Allows user to open a file.|
| **msoFileDialogSaveAs** . Allows user to save a file.|

## Example

In this example, Microsoft Excel opens the file dialog allowing the user to select one or more files. Once these files are selected, Excel displays the path for each file in a separate message.


```vb
Sub UseFileDialogOpen() 
 
    Dim lngCount As Long 
 
    ' Open the file dialog 
    With Application.FileDialog(msoFileDialogOpen) 
        .AllowMultiSelect = True 
        .Show 
 
        ' Display paths of each file selected 
        For lngCount = 1 To .SelectedItems.Count 
            MsgBox .SelectedItems(lngCount) 
        Next lngCount 
 
    End With 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

