---
title: CoAuthoring.CanShare Property (Word)
keywords: vbawd10.chm254869512
f1_keywords:
- vbawd10.chm254869512
ms.prod: word
api_name:
- Word.CoAuthoring.CanShare
ms.assetid: 9b0a08f8-cc54-5017-a487-bfab4057b711
ms.date: 06/08/2017
---


# CoAuthoring.CanShare Property (Word)

Returns a  **Boolean** that specifies whether this document can be co authored. Read-only.


## Syntax

 _expression_ . **CanShare**

 _expression_ An expression that returns a **[CoAuthoring](coauthoring-object-word.md)** object.


## Remarks

The value of this property is affected by whether  **[CanMerge](coauthoring-canmerge-property-word.md)** is **True** , the file extension is .docx, and the document is stored on a server that supports the File Synchronization via SOAP over HTTP protocol.


## Example

The following code example displays whether the active document can be co authored.


```vb
If ActiveDocument.CoAuthoring.CanShare Then 
    MsgBox "This document can be co authored." 
Else: MsgBox "This document cannot be co authored." 
End If
```


## See also


#### Concepts


[CoAuthoring Object](coauthoring-object-word.md)

