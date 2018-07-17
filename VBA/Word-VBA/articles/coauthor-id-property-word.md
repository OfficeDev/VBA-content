---
title: CoAuthor.ID Property (Word)
keywords: vbawd10.chm81068033
f1_keywords:
- vbawd10.chm81068033
ms.prod: word
api_name:
- Word.CoAuthor.ID
ms.assetid: a3118c4d-c4c7-9084-3182-8a449f32b020
ms.date: 06/08/2017
---


# CoAuthor.ID Property (Word)

Returns a  **String** that specifies a unique identifier for the specified author. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ An expression that returns a **CoAuthor** object.


## Remarks

The unique identifier returned by the  **ID** property should not be assumed to have a particular length or format.


## Example

The following code example displays the unique identifier for each co author in the active document.


```vb
Dim allAuthors As CoAuthors 
Dim coAuth As CoAuthor 
 
Set allAuthors = ActiveDocument.CoAuthoring.Authors 
 
For Each coAuth In allAuthors 
 MsgBox "The ID for " &; _ 
 coAuth.Name &; " is " &; coAuth.ID &; "." 
Next coAuth
```


## See also


#### Concepts


[CoAuthor Object](coauthor-object-word.md)

