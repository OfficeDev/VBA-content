---
title: Range.GoToEditableRange Method (Word)
keywords: vbawd10.chm157155743
f1_keywords:
- vbawd10.chm157155743
ms.prod: word
api_name:
- Word.Range.GoToEditableRange
ms.assetid: 4901bcef-56a7-c00e-409e-da0d442344c6
ms.date: 06/08/2017
---


# Range.GoToEditableRange Method (Word)

Returns a  **Range** object that represents an area of a document that can be modified by the specified user or group of users.


## Syntax

 _expression_ . **GoToEditableRange**( **_EditorID_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

You can also use the  **NextRange** property of the **Editor** object to return the next range for which the user has permission to modify.


## Example

The following example goes to the next range for which the current user has permission to modify.


```
Selection.GoToEditableRange wdEditorCurrent
```


## See also


#### Concepts


[Range Object](range-object-word.md)

