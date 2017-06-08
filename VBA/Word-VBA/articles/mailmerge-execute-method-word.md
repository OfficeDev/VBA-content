---
title: MailMerge.Execute Method (Word)
keywords: vbawd10.chm153092201
f1_keywords:
- vbawd10.chm153092201
ms.prod: word
api_name:
- Word.MailMerge.Execute
ms.assetid: ffce766a-2e2d-9633-e1d8-129a3976cadd
ms.date: 06/08/2017
---


# MailMerge.Execute Method (Word)

Performs the specified mail merge operation.


## Syntax

 _expression_ . **Execute**( **_Pause_** )

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pause_|Optional| **Variant**| **True** for Microsoft Word pause and display a troubleshooting dialog box if a mail merge error is found. **False** to report errors in a new document.|

## Example

This example executes a mail merge if the active document is a main document with an attached data source.


```vb
Set myMerge = ActiveDocument.MailMerge 
If myMerge.State = wdMainAndDataSource Then MyMerge.Execute
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

