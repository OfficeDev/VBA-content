---
title: SaveAs Method
keywords: vbagr10.chm3077082
f1_keywords:
- vbagr10.chm3077082
ms.prod: excel
api_name:
- Excel.SaveAs
ms.assetid: d8b3e963-e50a-3307-9abf-4ea37c46f114
ms.date: 06/08/2017
---


# SaveAs Method

Saves changes to the graph in a different file.

 _expression_. **SaveAs**( **_FileName_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **FileName**Required  **String**. A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.

## Example

This example creates a new workbook, prompts the user for a file name, and then saves the workbook.


```vb
Set NewBook = Workbooks.Add 
Do 
 fName = Application.GetSaveAsFilename 
Loop Until fName <> False 
NewBook.SaveAs Filename:=fName
```


