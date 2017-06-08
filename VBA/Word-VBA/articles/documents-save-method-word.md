---
title: Documents.Save Method (Word)
keywords: vbawd10.chm158072845
f1_keywords:
- vbawd10.chm158072845
ms.prod: word
api_name:
- Word.Documents.Save
ms.assetid: 547ba7a6-3ef5-10db-834d-58fc62502454
ms.date: 06/08/2017
---


# Documents.Save Method (Word)

Saves all the documents in the  **Documents** collection.


## Syntax

 _expression_ . **Save**( **_NoPrompt_** , **_OriginalFormat_** )

 _expression_ Required. A variable that represents a **[Documents](documents-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NoPrompt_|Optional| **Variant**| **True** to have Word automatically save all documents. **False** to have Word prompt the user to save each document that has changed since it was last saved.|
| _OriginalFormat_|Optional| **Variant**|Specifies the way the documents are saved. Can be one of the  **WdOriginalFormat** constants.|

## Remarks

If a document has not been saved before, the  **Save As** dialog box prompts the user for a file name.


## See also


#### Concepts


[Documents Collection Object](documents-object-word.md)

