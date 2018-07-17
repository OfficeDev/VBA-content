---
title: Range.Updates Property (Word)
keywords: vbawd10.chm157155833
f1_keywords:
- vbawd10.chm157155833
ms.prod: word
api_name:
- Word.Range.Updates
ms.assetid: 584c9a40-0975-75d9-e3d4-32e857fb62e5
ms.date: 06/08/2017
---


# Range.Updates Property (Word)

Returns a [CoAuthUpdates](http://msdn.microsoft.com/library/4a164415-0c6c-213b-da94-744e2394d1ef%28Office.15%29.aspx) collection object that represents all updates that were merged into the specified range at the last explicit save. Read-only.


## Syntax

 _expression_ . **Updates**

 _expression_ An expression that returns a **[Range](range-object-word.md)** object.


## Remarks

Use the  **Updates** property to return the[CoAuthUpdates](http://msdn.microsoft.com/library/4a164415-0c6c-213b-da94-744e2394d1ef%28Office.15%29.aspx) collection.


 **Note**  This property is only available for co authoring enabled documents. If you attempt to access this property on a document that is not enabled for co authoring, you will receive a run-time error.


## Example

The following code example displays the number of updates that were merged into the first paragraph of the active document at the last explicit save.


```vb
Dim countOfUpdates As Integer 
 
countOfUpdates = ActiveDocument.Paragraphs(1).Range.Updates.Count 
 
MsgBox "The number of updates is " &; countOfUpdates
```


## See also


#### Concepts


[Range Object](range-object-word.md)

