---
title: FindReplace.ReplaceWithText Property (Publisher)
keywords: vbapb10.chm8323077
f1_keywords:
- vbapb10.chm8323077
ms.prod: publisher
api_name:
- Publisher.FindReplace.ReplaceWithText
ms.assetid: 7bd0457f-c55e-3350-fe16-b9eac7d7d4fa
ms.date: 06/08/2017
---


# FindReplace.ReplaceWithText Property (Publisher)

Sets or retrieves a  **String** representing the replacement text in the specified range or selection. Read/write.


## Syntax

 _expression_. **ReplaceWithText**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

String


## Remarks

The default setting of the  **ReplaceWithText** property is an empty **String**.

If the  **ReplaceScope** property is set to either **pbReplaceScopeOne** or **pbReplaceScopeAll** and the **ReplaceWithText** property is not set, the text found will be replaced with the default empty string, thus removing the text.


## Example

The following example replaces all occurrences of the word "hello" with "goodbye" in the active document.


```vb
With ActiveDocument.Find 
 .Clear 
 .FindText = "hello" 
 .ReplaceWithText = "goodbye" 
 .MatchWholeWord = True 
 .ReplaceScope = pbReplaceScopeAll 
 .Execute 
End With
```


