---
title: Document.SelectAllEditableRanges Method (Word)
keywords: vbawd10.chm158007764
f1_keywords:
- vbawd10.chm158007764
ms.prod: word
api_name:
- Word.Document.SelectAllEditableRanges
ms.assetid: 510cd397-4c39-f36b-ed59-524247b35f16
ms.date: 06/08/2017
---


# Document.SelectAllEditableRanges Method (Word)

Selects all ranges for which the specified user or group of users has permission to modify.


## Syntax

 _expression_ . **SelectAllEditableRanges**( **_EditorID_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EditorID_|Optional| **Variant**|Can be either a  **String** that represents the user's e-mail alias (if in the same domain), an e-mail address, or a **WdEditorType** constant that represents a group of users. If omitted, only ranges for which all users have permissions will be selected.|

## Example

The following example selects all ranges for which the current user has permission to modify.


```vb
ActiveDocument.SelectAllEditableRanges wdEditorCurrent
```


## See also


#### Concepts


[Document Object](document-object-word.md)

