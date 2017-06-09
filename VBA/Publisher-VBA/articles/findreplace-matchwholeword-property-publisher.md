---
title: FindReplace.MatchWholeWord Property (Publisher)
keywords: vbapb10.chm8323083
f1_keywords:
- vbapb10.chm8323083
ms.prod: publisher
api_name:
- Publisher.FindReplace.MatchWholeWord
ms.assetid: 512d37bc-c900-ee17-8a8e-5875db0a2f85
ms.date: 06/08/2017
---


# FindReplace.MatchWholeWord Property (Publisher)

Sets or returns a  **Boolean** that represents whether the whole word will be matched in the search operation. Read/write. **Boolean**.


## Syntax

 _expression_. **MatchWholeWord**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

Boolean


## Remarks

The default value for  **MatchWholeWord** is **False**.


## Example

This example will select each occurrence of the word "fact" and apply bold formatting.


```vb
With ActiveDocument.Find 
 .Clear 
 .MatchWholeWord = True 
 .FindText = "fact" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With 

```

This example follows the previous example except that whole words will not be matched. Therefore the word "fact" within the word "factory" or "factoid" will also have bold formatting applied.




```vb
With ActiveDocument.Find 
 .Clear 
 .MatchWholeWord = False 
 .FindText = "fact" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With 

```


