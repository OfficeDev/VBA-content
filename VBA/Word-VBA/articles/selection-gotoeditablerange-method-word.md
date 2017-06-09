---
title: Selection.GoToEditableRange Method (Word)
keywords: vbawd10.chm158663683
f1_keywords:
- vbawd10.chm158663683
ms.prod: word
api_name:
- Word.Selection.GoToEditableRange
ms.assetid: 01c287a4-9293-22c1-9439-4a069a1e7299
ms.date: 06/08/2017
---


# Selection.GoToEditableRange Method (Word)

Returns a  **Range** object that represents an area of a document that can be modified by the specified user or group of users.


## Syntax

 _expression_ . **GoToEditableRange**( **_EditorID_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EditorID_|Optional| **Variant**|Can be either a  **String** that represents the user's e-mail alias (if in the same domain), an e-mail address, or a **WdEditorType** constant that represents a group of users. If omitted, selects all ranges for which all users have permissions to edit.|

### Return Value

Range


## Remarks

You can also use the  **[NextRange](editor-nextrange-property-word.md)** property of the **Editor** object to return the next range for which the user has permission to modify.


## Example

The following example goes to the next range for which the current user has permission to modify.


```
Selection.GoToEditableRange wdEditorCurrent
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

