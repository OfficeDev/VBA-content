---
title: Application.WorkbookAfterSave Event (Excel)
keywords: vbaxl10.chm504114
f1_keywords:
- vbaxl10.chm504114
ms.prod: excel
api_name:
- Excel.Application.WorkbookAfterSave
ms.date: 06/08/2017
---


# Application.WorkbookAfterSave Event (Excel)

Occurs after the workbook is saved.

**NOTE:** In Office 365, Excel supports AutoSave, which enables the user's edits to be saved automatically and continuously. Following the guidance in [this article](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md) will ensure that running code in response to the **WorkbookAfterSave** event will function as intended when AutoSave is enabled.

## Syntax

 _expression_ . **WorkbookAfterSave**( **_Wb_** , **_Success_** )

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook being saved.|
| _Success_|Required| **Boolean**|Returns  **True** if the save operation was successful; otherwise **False** .|

### Return Value

Nothing


## Example

The following code example displays a message box if the workbook was successfully saved. This code must be placed in a class module and an instance of that class must be correctly initialized. For more information about how to use event procedures with the  **Application** object, see [Using Events with the Application Object](http://msdn.microsoft.com/library/0063feba-47fd-29be-d2d5-8fcf47e70cbc%28Office.15%29.aspx).


```vb
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, _ 
 ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The " &; Wb.Name &; " workbook was successfully saved.") 
End If 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)