---
title: MailMergeDataSource.HeaderSourceType Property (Word)
keywords: vbawd10.chm152895492
f1_keywords:
- vbawd10.chm152895492
ms.prod: word
api_name:
- Word.MailMergeDataSource.HeaderSourceType
ms.assetid: e3ac1282-5f61-1425-07d7-d23a027decaf
ms.date: 06/08/2017
---


# MailMergeDataSource.HeaderSourceType Property (Word)

Returns a value that indicates the way the header source is being supplied for the mail merge operation. Read-only  **WdMailMergeDataSource** .


## Syntax

 _expression_ . **HeaderSourceType**

 _expression_ Required. A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks


 **Security Note**  




## Example

This example opens the header source attached to the active document if the source is a Word document.


```vb
Dim mmdsTemp As MailMergeDataSource 
 
Set mmdsTemp = ActiveDocument.MailMerge.DataSource 
 
If mmdsTemp.HeaderSourceType = wdMergeInfoFromWord Then 
 Documents.Open FileName:=mmdsTemp.HeaderSourceName 
End If
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)

