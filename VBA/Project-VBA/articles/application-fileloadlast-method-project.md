---
title: Application.FileLoadLast Method (Project)
keywords: vbapj.chm117
f1_keywords:
- vbapj.chm117
ms.prod: project-server
api_name:
- Project.Application.FileLoadLast
ms.assetid: c775d573-d184-d3ac-ed81-3552cc9b045b
ms.date: 06/08/2017
---


# Application.FileLoadLast Method (Project)

Opens one of the recently used files.


## Syntax

 _expression_. **FileLoadLast**( ** _Number_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Number_|Optional|**Integer**|A number that specifies which of the most recently used files to open. The maximum value is 17 in a default installation of Project.|

### Return Value

 **Boolean**


## Remarks

To specify the number of files to show on the  **Recent** tab of the Backstage view, change the value in the **Show this number of recent documents** drop-down list in the **Display** section of the **Advanced** tab of the **Project Options** dialog box. The maximum number possible is 50.


## Example

The following example opens the five most recently used files. It assumes the "Recently Used File List" option has been selected.


```vb
Sub OpenThe9MRUFiles() 
 
 Dim i As Integer ' Index used in For...Next loop 
 
 For i = 1 To 5 
 FileLoadLast i 
 ' Ignore errors that may be due to missing files. 
 On Error Resume Next 
 Next i 
 
End Sub
```


