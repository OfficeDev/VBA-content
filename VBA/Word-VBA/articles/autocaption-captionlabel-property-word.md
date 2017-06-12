---
title: AutoCaption.CaptionLabel Property (Word)
keywords: vbawd10.chm159055875
f1_keywords:
- vbawd10.chm159055875
ms.prod: word
api_name:
- Word.AutoCaption.CaptionLabel
ms.assetid: 8e4864e4-e42b-ccc0-9611-eda7753089f4
ms.date: 06/08/2017
---


# AutoCaption.CaptionLabel Property (Word)

Returns or sets the caption label ("Figure," "Table," or "Equation," for example) of the specified caption. Read/write  **Variant** .


## Syntax

 _expression_ . **CaptionLabel**

 _expression_ A variable that represents an **[AutoCaption](autocaption-object-word.md)** object.


## Remarks

This property can be set to a string or a  **WdCaptionLabelID** constant.


## Example

This example displays the name ("Microsoft Excel Worksheet," for example) and caption label ("Figure," for example) for each item that has a caption added automatically when inserted.


```vb
Dim acLoop As AutoCaption 
 
For Each acLoop In AutoCaptions 
 If acLoop.AutoInsert = True Then MsgBox acLoop.Name _ 
 &; vbCr &; "Label = " &; acLoop.CaptionLabel.Name 
Next acLoop
```

This example sets the caption label for Word tables to "Table" and then inserts a new table immediately after the selection.




```vb
With AutoCaptions("Microsoft Word Table") 
 .AutoInsert = True 
 .CaptionLabel = wdCaptionTable 
End With 
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=2, _ 
 NumColumns:=3
```


## See also


#### Concepts


[AutoCaption Object](autocaption-object-word.md)

