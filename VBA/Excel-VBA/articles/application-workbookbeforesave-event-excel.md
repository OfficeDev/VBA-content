---
title: Application.WorkbookBeforeSave Event (Excel)
keywords: vbaxl10.chm504085
f1_keywords:
- vbaxl10.chm504085
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforeSave
ms.assetid: e93a7cef-b018-ddab-c96f-b3215143f31f
ms.date: 06/08/2017
---


# Application.WorkbookBeforeSave Event (Excel)

Occurs before any open workbook is saved.

**NOTE:** In Office 365, Excel supports AutoSave, which enables the user's edits to be saved automatically and continuously. Following the guidance in [this article](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md) will ensure that running code in response to the **WorkbookBeforeSave** event will function as intended when AutoSave is enabled.

## Syntax

 _expression_ . **WorkbookBeforeSave**( **_Wb_** , **_SaveAsUI_** , **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook.|
| _SaveAsUI_|Required| **Boolean**| **True** if the **Save As** dialog box will be displayed due to changes made that need to be saved in the workbook.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the workbook isn't saved when the procedure is finished.|

### Return Value

Nothing


## Example

This example prompts the user for a yes or no response before saving any workbook.


```vb
Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, _ 
 ByVal SaveAsUI As Boolean, Cancel as Boolean) 
 a = MsgBox("Do you really want to save the workbook?", vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)