---
title: Comment.Previous Method (Excel)
keywords: vbaxl10.chm516079
f1_keywords:
- vbaxl10.chm516079
ms.prod: excel
api_name:
- Excel.Comment.Previous
ms.assetid: b7854b0f-0e88-6749-2e62-6d45add8b945
ms.date: 06/08/2017
---


# Comment.Previous Method (Excel)

Returns a  **[Comment](comment-object-excel.md)** object that represents the previous comment.


## Syntax

 _expression_ . **Previous**

 _expression_ An expression that returns a **Comment** object.


### Return Value

Comment


## Remarks

This method works only on one sheet. Using this method on the first comment on a sheet returns  **Null** (not the last comment on the previous sheet).


## Example

This example deletes every second comment, navigating with the  **Previous** method.


 **Note**  Please test in a new workbook with no existing comments. To clear all comments from a workbook use  `Selection.SpecialCells(xlCellTypeComments).Delete` in the **Immediate Pane** .


```vb
'Sets up the comments 
For xNum = 1 To 10 
 Range("A" &; xNum).AddComment 
 Range("A" &; xNum).Comment.Text Text:="Comment " &; xNum 
Next 
 
MsgBox "Comments created... A1:A10" 
 
'Deletes every second comment in the A1:A10 range 
For yNum = 10 To 1 Step -2 
 Range("A" &; yNum).Comment.Previous.Shape.Select True 
 Selection.Delete 
Next 
 
MsgBox "Deleted every second comment"
```


## See also


#### Concepts


[Comment Object](comment-object-excel.md)

