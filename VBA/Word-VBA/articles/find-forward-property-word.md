---
title: Find.Forward Property (Word)
keywords: vbawd10.chm162529290
f1_keywords:
- vbawd10.chm162529290
ms.prod: word
api_name:
- Word.Find.Forward
ms.assetid: deacedde-ca81-6fa0-6a62-696163d8c52d
ms.date: 06/08/2017
---


# Find.Forward Property (Word)

 **True** if the find operation searches forward through the document. Read/write **Boolean** .


## Syntax

 _expression_ . **Forward**

 _expression_ A variable that represents a **[Find](find-object-word.md)** object.


## Remarks

 **False** causes a Word find operation to search backward through the document.


## Example

This example replaces the next occurrence of the word "hi" in the selection with "hello."


```vb
With Selection.Find 
 .Forward = True 
 .Text = "hi" 
 .ClearFormatting 
 .Replacement.Text = "hello" 
 .Execute Replace:=wdReplaceOne 
End With
```

The following example searches backward through the document for the word "Microsoft." If the word is found, it is automatically selected.




```vb
Selection.Collapse Direction:=wdCollapseStart 
With Selection.Find 
 .Forward = False 
 .ClearFormatting 
 .MatchWholeWord = True 
 .MatchCase = False 
 .Wrap = wdFindContinue 
 .Execute FindText:="Microsoft" 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

