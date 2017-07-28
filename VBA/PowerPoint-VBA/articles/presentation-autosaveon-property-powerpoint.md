---
title: Presentation.AutoSaveOn Property (PowerPoint)
keywords: vbapp10.chm583129
f1_keywords:
- vbapp10.chm583129
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.AutoSaveOn
ms.date: 07/28/2017
---


# Presentation.AutoSaveOn Property (PowerPoint)

**True** if the edits in the presentation are automatically saved. Read/write **Boolean**.

## Syntax

_expression_.**AutoSaveOn**

_expression_ A variable that represents a **Presentation** object.

## Remarks

When a new presentation is created, the default value for the **AutoSaveOn** property is **False**, the property is disabled, and the user's changes will need to be saved manually. However, if the presentation is hosted on the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online), then the **AutoSaveOn** property defaults to **True** and the edits in the specified presentation are automatically saved. If a cloud-hosted presentation is shared with other users, then their changes will also be automatically merged into the user's local copy when **AutoSaveOn** is **True**.

**Table 1 AutoSaveOn behavior**

|`AutoSaveOn` Toggle State|Set `AutoSaveOn` to True|Set `AutoSaveOn` to False|
|:-----|:-----|:-----|
|`AutoSaveOn == True`|No-op|`AutoSaveOn` turned off|
|`AutoSaveOn == False`|`AutoSaveOn` turned on|No-op|
|Disabled|Error|No-op|

## Example

This example notifies you whether the presentation is set to be automatically saved or not.

```vb
Sub UseAutoSaveOn()
    MsgBox "This presentation is being saved automatically: " & ActivePresentation.AutoSaveOn
End Sub
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)

[Presentation Object](presentation-object-powerpoint.md)
