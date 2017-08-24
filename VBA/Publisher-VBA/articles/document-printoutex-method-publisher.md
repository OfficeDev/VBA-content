---
title: Document.PrintOutEx Method (Publisher)
keywords: vbapb10.chm196755
f1_keywords:
- vbapb10.chm196755
ms.prod: publisher
api_name:
- Publisher.Document.PrintOutEx
ms.assetid: f11b6f8b-08a0-28f6-5930-47d684585bef
ms.date: 06/08/2017
---


# Document.PrintOutEx Method (Publisher)

Prints all or part of the specified publication.


## Syntax

 _expression_. **PrintOut**( **_From_**,  **_To_**,  **_PrintToFile_**,  **_Copies_**,  **_Collate_**,  **_PrintStyle_**)

 _expression_A variable that represents a  **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|From|Optional| **Long**|The starting page number.|
|To|Optional| **Long**|The ending page number.|
|PrintToFile|Optional| **String**|The path and file name of a document to be printed to a file.|
|Copies|Optional| **Long**|The number of copies to be printed.|
|Collate|Optional| **Boolean**|When printing multiple copies of a document,  **True** to print all pages of the document before printing the next copy.|
|PrintStyle|Optional| **PbPrintStyle**|The print style to use. See Remarks for possible values.|

## Remarks

The PrintStyle parameter can be one of the  **[PbPrintStyle](pbprintstyle-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.

If PrintStyle is  **pbPrintStyleMultipleCopiesPerSheet** or **pbPrintStyleMultiplePagesPerSheet**, Publisher ignores any value you pass for the Collate parameter.


## Example

This example prints the active publication.


```vb
Sub PrintActivePublication() 
 ThisDocument.PrintOutEx 
End Sub
```


