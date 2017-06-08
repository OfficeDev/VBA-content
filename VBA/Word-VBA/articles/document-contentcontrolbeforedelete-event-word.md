---
title: Document.ContentControlBeforeDelete Event (Word)
keywords: vbawd10.chm4001011
f1_keywords:
- vbawd10.chm4001011
ms.prod: word
api_name:
- Word.Document.ContentControlBeforeDelete
ms.assetid: a690fb97-0de3-de0e-7e84-edaaea756e83
ms.date: 06/08/2017
---


# Document.ContentControlBeforeDelete Event (Word)

Occurs before removing a content control from a document.


## Syntax

Private Sub  _expression_ _**ContentControlBeforeDelete**( **_OldContentControl_** , **_InUndoRedo_** )

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OldContentControl_|Required| **ContentControl**|The content control being deleted.|
| _InUndoRedo_|Required| **Boolean**| Specifies whether the removal is taking place as part an undo or redo action.|

## Remarks

For information about using events with the  **Document** object, see[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


## See also


#### Concepts


[Document Object](document-object-word.md)

