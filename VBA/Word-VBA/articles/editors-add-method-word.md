---
title: Editors.Add Method (Word)
keywords: vbawd10.chm9175541
f1_keywords:
- vbawd10.chm9175541
ms.prod: word
api_name:
- Word.Editors.Add
ms.assetid: d17ad2dc-1607-6cb3-f7e4-eefcd7fc3202
ms.date: 06/08/2017
---


# Editors.Add Method (Word)

Returns an  **Editor** object that represents a new permission for a specified user to modify a range or selection within a document. .


## Syntax

 _expression_ . **Add**( **_EditorID_** )

 _expression_ Required. A variable that represents an **[Editors](editors-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EditorID_|Optional| **Variant**|Can be either a  **String** that represents the user's e-mail alias (if in the same domain), an e-mail address, or a **WdEditorType** that represents a group of users.|

## Example

The following example gives editing permissions for the selected text to the current user.


```vb
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors.Add(wdEditorCurrent)
```


## See also


#### Concepts


[Editors Collection](editors-object-word.md)

