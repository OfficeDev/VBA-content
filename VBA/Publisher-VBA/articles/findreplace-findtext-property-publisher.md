---
title: FindReplace.FindText Property (Publisher)
keywords: vbapb10.chm8323076
f1_keywords:
- vbapb10.chm8323076
ms.prod: publisher
api_name:
- Publisher.FindReplace.FindText
ms.assetid: 5c8d2803-174e-a82f-d94c-3d96c4b4a2eb
ms.date: 06/08/2017
---


# FindReplace.FindText Property (Publisher)

Sets or retrieves a  **String** representing the text to find in the specified range or selection. Read/write.


## Syntax

 _expression_. **FindText**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

String


## Remarks

The  **FindText** property returns the plain, unformatted text of the selection. When you set this property, the search text is specified. You can search for special characters by specifying appropriate character codes. For example, "^p" corresponds to a paragraph mark and "^t" corresponds to a tab character.

The default value for the  **FindText** property is an empty string. Because only text searching is supported, **FindText** must be explicitly set to avoid a runtime error.


## Example

This example replaces all occurrences of the word "This" in the selection with "That" in each open publication.


```vb
Dim objDocument As Document 
 
For Each objDocument In Documents 
 With objDocument.Find 
 .Clear 
 .MatchCase = True 
 .FindText = "This" 
 .ReplaceWithText = "That" 
 .ReplaceScope = pbReplaceScopeAll 
 .Forward = True 
 .Execute 
 End With 
Next objDocument 

```


