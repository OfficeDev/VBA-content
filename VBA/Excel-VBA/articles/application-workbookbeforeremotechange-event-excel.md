---
title: Application.WorkbookBeforeRemoteChange Event (Excel)
keywords: vbaxl10.chm503113
f1_keywords:
- vbaxl10.chm503113
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforeRemoteChange
ms.date: 06/08/2017
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

This example shows you where you can place code that runs before merging an incoming remote change. This code must be placed in a class module and an instance of that class must be correctly initialized. For more information about how to use event procedures with the  **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object.md).

```vb
Private Sub App_WorkbookBeforeRemoteChange(ByVal Wb As Workbook)
    'A remote user has made a change to this workbook.
    'The code in this subroutine will be run before those changes are merged.
End Sub
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)

[Co authoring](about-coauthoring-in-excel.md)

[Workbook Object](workbook-object-excel.md)
