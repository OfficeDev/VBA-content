---
title: MailMergeDataSource.QueryString Property (Word)
keywords: vbawd10.chm152895494
f1_keywords:
- vbawd10.chm152895494
ms.prod: word
api_name:
- Word.MailMergeDataSource.QueryString
ms.assetid: 8b2d7490-d3f1-bc46-043f-f37fb2e2fa91
ms.date: 06/08/2017
---


# MailMergeDataSource.QueryString Property (Word)

Returns or sets the query string (SQL statement) used to retrieve a subset of the data in a mail merge data source. Read/write  **String** .


## Syntax

 _expression_ . **QueryString**

 _expression_ An expression that returns a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Example

This example returns the query string for the data source attached to the active document.


```
qString = ActiveDocument.MailMerge.DataSource.QueryString
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)

