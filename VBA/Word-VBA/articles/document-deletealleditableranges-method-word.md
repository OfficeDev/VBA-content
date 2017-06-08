---
title: Document.DeleteAllEditableRanges Method (Word)
keywords: vbawd10.chm158007765
f1_keywords:
- vbawd10.chm158007765
ms.prod: word
api_name:
- Word.Document.DeleteAllEditableRanges
ms.assetid: 021456eb-516c-5616-3e32-19d0b9908aef
ms.date: 06/08/2017
---


# Document.DeleteAllEditableRanges Method (Word)

Deletes permissions in all ranges for which the specified user or group of users has permission to modify.


## Syntax

 _expression_ . **DeleteAllEditableRanges**( **_EditorID_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EditorID_|Optional| **Variant**|Can be either a  **String** that represents the user's e-mail alias (if in the same domain), an e-mail address, or a **WdEditorType** constant that represents a group of users. If omitted, no permissions are deleted from a document.|

## Remarks

You can also use the  **[DeleteAll](editor-deleteall-method-word.md)** method to delete permissions in all ranges for which a specified user or group of users has permission to modify.


## Example

The following example deletes all permissions in all ranges for the current user.


```vb
ActiveDocument.DeleteAllEditableRanges wdEditorCurrent
```


## See also


#### Concepts


[Document Object](document-object-word.md)

