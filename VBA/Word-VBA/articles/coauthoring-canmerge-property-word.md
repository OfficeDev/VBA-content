---
title: CoAuthoring.CanMerge Property (Word)
keywords: vbawd10.chm254869513
f1_keywords:
- vbawd10.chm254869513
ms.prod: word
api_name:
- Word.CoAuthoring.CanMerge
ms.assetid: c74efdfe-9f8b-e524-14fb-7866ae0e34ae
ms.date: 06/08/2017
---


# CoAuthoring.CanMerge Property (Word)

Returns a  **Boolean** that specifies whether the document can be auto-merged. Read-only.


## Syntax

 _expression_ . **CanMerge**

 _expression_ An expression that returns a **[CoAuthoring](coauthoring-object-word.md)** object.


## Remarks

Only documents stored on a server that supports the File Synchronization via SOAP over HTTP protocol can be co authored, for example, SharePoint Server 2010. Additionally, a document that has the following features cannot be auto-merged:


- Digital Rights Management
    
- Digital Signatures
    
- Final mode
    
- Encryption
    
- Master Document role
    
- Disable Auto-merge Client Policy
    
- Framesets
    
- Object Linking and Embedding (OLE) objects that do not have Revision Save IDs (RSIDs)
    
- ActiveX controls
    
- OfficeArt Engine 2.0 Charts and ink objects, and SmartArt that do not have corresponding IDs in the document
    
- Documents with file name extensions other than .docx, .doc, and .odt
    
- Word Blog documents
    



## Example

The following code example displays whether the active document can be auto-merged.


```vb
If ActiveDocument.CoAuthoring.CanMerge Then 
    MsgBox "This document can be auto-merged." 
Else: MsgBox "This document cannot be auto-merged." 
End If
```


## See also


#### Concepts


[CoAuthoring Object](coauthoring-object-word.md)

