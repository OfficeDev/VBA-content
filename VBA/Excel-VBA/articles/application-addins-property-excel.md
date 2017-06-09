---
title: Application.AddIns Property (Excel)
keywords: vbaxl10.chm132081
f1_keywords:
- vbaxl10.chm132081
ms.prod: excel
api_name:
- Excel.Application.AddIns
ms.assetid: 0798690a-910a-b832-e143-df51d7c061ca
ms.date: 06/08/2017
---


# Application.AddIns Property (Excel)

Returns an  **[AddIns](addins-object-excel.md)** collection that represents all the add-ins listed in the **Add-Ins** dialog box ( **Add-Ins** command on the **Developer** tab). Read-only.


## Syntax

 _expression_ . **AddIns**

 _expression_ A variable that represents an **Application** object.


## Remarks

Using this method without an object qualifier is equivalent to  `Application.Addins`.


## Example

This example displays the status of the Analysis ToolPak add-in. Note that the string used as the index to the  **AddIns** collection is the title of the add-in, not the add-in's file name.


```vb
If AddIns("Analysis ToolPak").Installed = True Then 
 MsgBox "Analysis ToolPak add-in is installed" 
Else 
 MsgBox "Analysis ToolPak add-in is not installed" 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

