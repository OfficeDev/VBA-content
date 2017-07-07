---
title: Workbook.autoSaveOn Property (Excel)
keywords: vbaxl10.chm199287
f1_keywords:
- vbaxl10.chm199287
ms.prod: excel
api_name:
- Excel.Workbook.autoSaveOn
ms.date: 06/08/2017
---


# Workbook.autoSaveOn Property (Excel)

**True** if the edits in the workbook are automatically saved. Read/write **Boolean**.

## Syntax

_expression_.**autoSaveOn**

_expression_ A variable that represents a **Workbook** object.

## Remarks

When a new workbook is created, the default value for the **autoSaveOn** property is **False** and the user's changes will need to be saved manually. However, if the workbook is hosted on the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online), then the **autoSaveOn** property defaults to **True** and the edits in the specified workbook are automatically saved. If a cloud-hosted workbook is shared with other users, then their changes will also be automatically merged into the user's local copy when **autoSaveOn** is **True**.

## Example

This example notifies you whether the workbook is set to be automatically saved or not.

```vb
Sub UseAutoSaveOn()
    ActiveWorkbook.autoSaveOn = True
    MsgBox "This workbook is being saved automatically: " & ActiveWorkbook.autoSaveOn
End Sub
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)

[co authoring](about-coauthoring-in-excel.md)

[Workbook Object](workbook-object-excel.md)
