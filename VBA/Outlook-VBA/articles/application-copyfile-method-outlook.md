---
title: Application.CopyFile Method (Outlook)
keywords: vbaol11.chm727
f1_keywords:
- vbaol11.chm727
ms.prod: outlook
api_name:
- Outlook.Application.CopyFile
ms.assetid: dc848d48-23e0-d0a9-049d-b2ae414151d5
ms.date: 06/08/2017
---


# Application.CopyFile Method (Outlook)

Copies a file from a specified location into a Microsoft Outlook store.


## Syntax

 _expression_ . **CopyFile**( **_FilePath_** , **_DestFolderPath_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FilePath_|Required| **String**|The path name of the object you want to copy.|
| _DestFolderPath_|Required| **String**|The location you want to copy the file to.|

### Return Value

An  **Object** value that represents the copied file.


## Example

The following Visual Basic for Applications (VBA) example creates a Microsoft Excel worksheet called 'MyExcelDoc.xlsx' and then copies it from the user's hard drive to the user's  **Inbox**. 


```vb
Sub CopyFileSample() 
 
 Dim strPath As String 
 
 Dim ExcelApp As Object 
 
 Dim ExcelSheet As Object 
 
 Dim doc As Object 
 
 
 
 
 
 Set ExcelApp = CreateObject("Excel.Application") 
 
 strPath = ExcelApp.DefaultFilePath &; "\MyExcelDoc.xlsx" 
 
 Set ExcelSheet = ExcelApp.Workbooks.Add 
 
 ExcelSheet.ActiveSheet.cells(1, 1).Value = 10 
 
 ExcelSheet.SaveAs strPath 
 
 ExcelApp.Quit 
 
 Set ExcelApp = Nothing 
 
 Set doc = Application.CopyFile(strPath, "Inbox") 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

