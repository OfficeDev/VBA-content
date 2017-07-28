---
title: Workbook.AutoSaveOn Property (Excel)
keywords: vbaxl10.chm199287
f1_keywords:
- vbaxl10.chm199287
ms.prod: excel
api_name:
- Excel.Workbook.AutoSaveOn
ms.date: 07/28/2017
---


# Workbook.AutoSaveOn Property (Excel)

**True** if the edits in the workbook are automatically saved. Read/write **Boolean**.

## Syntax

_expression_.**AutoSaveOn**

_expression_ A variable that represents a **Workbook** object.

## Remarks

When a new workbook is created, the default value for the **AutoSaveOn** property is **False**, the property is disabled, and the user's changes will need to be saved manually. However, if the workbook is hosted on the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online), then the **AutoSaveOn** property defaults to **True** and the edits in the specified workbook are automatically saved. If a cloud-hosted workbook is shared with other users, then their changes will also be automatically merged into the user's local copy when **AutoSaveOn** is **True**.

**Table 1 AutoSaveOn behavior**

|`AutoSaveOn` Toggle State|Set `AutoSaveOn` to True|Set `AutoSaveOn` to False|
|:-----|:-----|:-----|
|`AutoSaveOn == True`|No-op|`AutoSaveOn` turned off|
|`AutoSaveOn == False`|`AutoSaveOn` turned on|No-op|
|Disabled|Error|Error|

## Example

This example notifies you whether the workbook is set to be automatically saved or not.

```vb
Sub UseAutoSaveOn()
    MsgBox "This workbook is being saved automatically: " & ActiveWorkbook.AutoSaveOn
End Sub
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)

[Co authoring](about-coauthoring-in-excel.md)

[Workbook Object](workbook-object-excel.md)
