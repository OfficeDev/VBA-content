---
title: Document.SetCompatibilityMode Method (Word)
keywords: vbawd10.chm158007867
f1_keywords:
- vbawd10.chm158007867
ms.prod: word
api_name:
- Word.SetCompatibilityMode
ms.assetid: f167a640-340e-56ed-34c0-0c3dbff8575a
ms.date: 06/08/2017
---


# Document.SetCompatibilityMode Method (Word)

Sets the compatibility mode for the document.


## Syntax

 _expression_ . **SetCompatibilityMode**( **_Mode_** )

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Mode_|Required| **Long**|Specifies which version of Word to approximate. Use a constant from the [WdCompatibilityMode](wdcompatibilitymode-enumeration-word.md) enumeration as an argument for this parameter.|

## Remarks

When you open a document in Word that was created in a previous version of Word, Compatibility Mode is turned on. Compatibility Mode ensures that no new or enhanced features in Word are available while working with a document, so that people who edit the document using previous versions of Word will have full editing capabilities.


## Example

The following code example puts Word in Word 2003 Compatibility Mode.


```vb
ActiveDocument.SetCompatibilityMode (wdWord2003)
```


## See also


#### Concepts


[Document Object](document-object-word.md)

