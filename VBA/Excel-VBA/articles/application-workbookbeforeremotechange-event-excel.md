---
title: Application.WorkbookBeforeRemoteChange Event (Excel)
keywords: vbaxl10.chm503113
f1_keywords:
- vbaxl10.chm503113
ms.prod: EXCEL
api_name:
- Excel.Application.WorkbookBeforeRemoteChange
ms.assetid: application-workbookbeforeremotechange-event-excel
---


# Application.WorkbookBeforeRemoteChange Event (Excel)

Occurs before a remote user's edits to the workbook are merged.

## Syntax

 _expression_.**WorkbookBeforeRemoteChange**( **_Wb_** )

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Parameters

|**Name**|**Required or Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook which has been changed by a remote user.|

## Return value

Nothing

## Example

This example shows you where you can place code that runs before merging an incoming remote change. This code must be placed in a class module and an instance of that class must be correctly initialized. For more information about how to use event procedures with the  **Application** object, see [Using Events with the Application Object](http://msdn.microsoft.com/library/0063feba-47fd-29be-d2d5-8fcf47e70cbc%28Office.15%29.aspx).

```vb
Private Sub App_WorkbookBeforeRemoteChange(ByVal Wb As Workbook)
    'A remote user has made a change to this workbook.
    'The code in this subroutine will be run before those changes are merged.
End Sub
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/about-autosave.md)

[Co-authoring](about-coauthoring-in-excel.md)

[Workbook Object](workbook-object-excel.md)
