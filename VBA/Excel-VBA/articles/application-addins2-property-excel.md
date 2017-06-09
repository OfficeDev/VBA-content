---
title: Application.AddIns2 Property (Excel)
keywords: vbaxl10.chm133322
f1_keywords:
- vbaxl10.chm133322
ms.prod: excel
api_name:
- Excel.Application.AddIns2
ms.assetid: 3fd3de81-beae-c5b0-572d-c3f81e251db2
ms.date: 06/08/2017
---


# Application.AddIns2 Property (Excel)

Returns an  **[AddIns2](addins2-object-excel.md)** collection that represents all the add-ins that are currently available or open in Microsoft Excel, regardless of whether they are installed. Read-only


## Syntax

 _expression_ . **AddIns2**

 _expression_ A variable that returns an **Application** object.


## Example

This example displays the status of the Analysis ToolPak add-in. Note that the string used as the index to the  **AddIns** collection is the title of the add-in, not the add-in's file name.


```vb
If Application.AddIns2("Analysis ToolPak").Installed = True Then 
 MsgBox "Analysis ToolPak add-in is installed" 
Else 
 MsgBox "Analysis ToolPak add-in is not installed" 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

