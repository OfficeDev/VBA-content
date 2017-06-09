---
title: Workbook.EnableAutoRecover Property (Excel)
keywords: vbaxl10.chm199201
f1_keywords:
- vbaxl10.chm199201
ms.prod: excel
api_name:
- Excel.Workbook.EnableAutoRecover
ms.assetid: 04a82e4d-0231-adf1-1289-35514372c995
ms.date: 06/08/2017
---


# Workbook.EnableAutoRecover Property (Excel)

Saves changed files, of all formats, on a timed interval. Read/write  **Boolean** .


## Syntax

 _expression_ . **EnableAutoRecover**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

If Microsoft Excel fails, the system fails, or if the system is improperly shut down (not allowing Excel to save the changed files), the backed up files are opened and the user has an opportunity to save changes that otherwise would have been lost. When the user restarts Excel, a document recovery window opens, giving the user an option to recover the files they were working on. Setting this property to  **True** (default) enables this feature.


## Example

The following example checks the setting of the AutoRecover feature and if not enabled, Excel enables it and then notifies the user.


```vb
Sub UseAutoRecover() 
 
 ' Check to see if the feature is enabled, if not, enable it. 
 If ActiveWorkbook.EnableAutoRecover = False Then 
 ActiveWorkbook.EnableAutoRecover = True 
 MsgBox "The AutoRecover feature has been enabled." 
 Else 
 MsgBox "The AutoRecover feature is already enabled." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

