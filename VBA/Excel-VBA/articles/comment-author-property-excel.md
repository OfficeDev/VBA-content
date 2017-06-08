---
title: Comment.Author Property (Excel)
keywords: vbaxl10.chm516073
f1_keywords:
- vbaxl10.chm516073
ms.prod: excel
api_name:
- Excel.Comment.Author
ms.assetid: ac964a80-1646-41a0-8b3a-941c800395e7
ms.date: 06/08/2017
---


# Comment.Author Property (Excel)

Returns or sets the author of the comment. Read-only  **String** .


## Syntax

 _expression_ . **Author**

 _expression_ A variable that represents a **Comment** object.


## Example

This example deletes all comments added by Jean Selva on the active sheet.


```vb
For Each c in ActiveSheet.Comments 
 If c.Author = "Jean Selva" Then c.Delete 
Next
```


## See also


#### Concepts


[Comment Object](comment-object-excel.md)

