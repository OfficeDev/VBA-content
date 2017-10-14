---
title: Document.AutoSaveOn Property (Word)
keywords: vbawd10.chm158007920
f1_keywords:
- vbawd10.chm158007920
ms.prod: word
api_name:
- Word.Document.AutoSaveOn
ms.date: 07/28/2017
---


# Document.AutoSaveOn Property (Word)

**True** if the edits in the document are automatically saved. Read/write **Boolean**.

## Syntax

_expression_.**AutoSaveOn**

_expression_ A variable that represents a **Document** object.

## Remarks

When a new document is created, the default value for the **AutoSaveOn** property is **False**, the property is disabled, and the user's changes will need to be saved manually. However, if the document is hosted on the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online), then the **AutoSaveOn** property defaults to **True** and the edits in the specified document are automatically saved. If a cloud-hosted document is shared with other users, then their changes will also be automatically merged into the user's local copy when **AutoSaveOn** is **True**.

**Table 1 AutoSaveOn behavior**

|`AutoSaveOn` Toggle State|Set `AutoSaveOn` to True|Set `AutoSaveOn` to False|
|:-----|:-----|:-----|
|`AutoSaveOn == True`|No-op|`AutoSaveOn` turned off|
|`AutoSaveOn == False`|`AutoSaveOn` turned on|No-op|
|Disabled|Error|Error|

## Example

This example notifies you whether the document is set to be automatically saved or not.

```vb
Sub UseAutoSaveOn()
    MsgBox "This document is being saved automatically: " & ActiveDocument.AutoSaveOn
End Sub
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)

[Document Object](document-object-word.md)
