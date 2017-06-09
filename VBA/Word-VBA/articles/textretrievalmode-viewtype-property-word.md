---
title: TextRetrievalMode.ViewType Property (Word)
keywords: vbawd10.chm154730496
f1_keywords:
- vbawd10.chm154730496
ms.prod: word
api_name:
- Word.TextRetrievalMode.ViewType
ms.assetid: 1dbc3f48-6d99-84f4-b9db-73a25e8f07c0
ms.date: 06/08/2017
---


# TextRetrievalMode.ViewType Property (Word)

Returns or sets the view for the  **TextRetrievalMode** object. Read/write **WdViewType** .


## Syntax

 _expression_ . **ViewType**

 _expression_ Required. A variable that represents a **[TextRetrievalMode](textretrievalmode-object-word.md)** object.


## Remarks

Changing the view for the  **TextRetrievalMode** object doesn't change the display of a document on the screen. Instead, it determines which characters in the document will be included when a range is retrieved.


## Example

This example sets the view for text retrieval to outline view and then displays the contents of the active document in a dialog box. Note that only the text displayed in outline view is retrieved.


```vb
Set myText = ActiveDocument.Content 
myText.TextRetrievalMode.ViewType = wdOutlineView 
Msgbox myText
```


## See also


#### Concepts


[TextRetrievalMode Object](textretrievalmode-object-word.md)

