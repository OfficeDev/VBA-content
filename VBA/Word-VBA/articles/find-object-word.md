---
title: Find Object (Word)
keywords: vbawd10.chm2480
f1_keywords:
- vbawd10.chm2480
ms.prod: word
api_name:
- Word.Find
ms.assetid: da822788-cad5-992a-a835-18cc574cc324
ms.date: 06/08/2017
---


# Find Object (Word)

Represents the criteria for a find operation. 


## Remarks

The properties and methods of the  **Find** object correspond to the options in the **Find and Replace** dialog box.

Use the  **Find** property to return a **Find** object. The following example finds and selects the next occurrence of the word "hi."




```vb
With Selection.Find 
 .ClearFormatting 
 .Text = "hi" 
 .Execute Forward:=True 
End With
```

The following example finds all occurrences of the word "hi" in the active document and replaces the word with "hello."




```vb
Set myRange = ActiveDocument.Content 
myRange.Find.Execute FindText:="hi", ReplaceWith:="hello", _ 
 Replace:=wdReplaceAll
```

Remarks

If you've gotten to the  **Find** object from the **Selection** object, the selection is changed when text matching the find criteria is found. The following example selects the next occurrence of the word "blue."




```
Selection.Find.Execute FindText:="blue", Forward:=True
```

If you've gotten to the  **Find** object from the **Range** object, the selection isn't changed when text matching the find criteria is found, but the **Range** object is redefined. The following example locates the first occurrence of the word "blue" in the active document. If "blue" is found in the document, _myRange_ is redefined and bold formatting is applied to "blue."




```vb
Set myRange = ActiveDocument.Content 
myRange.Find.Execute FindText:="blue", Forward:=True 
If myRange.Find.Found = True Then myRange.Bold = True
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


