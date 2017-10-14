---
title: Document.XMLSaveThroughXSLT Property (Word)
keywords: vbawd10.chm158007771
f1_keywords:
- vbawd10.chm158007771
ms.prod: word
api_name:
- Word.Document.XMLSaveThroughXSLT
ms.assetid: cc25a073-99c5-f31b-0cad-b6e4c9a7ff0c
ms.date: 06/08/2017
---


# Document.XMLSaveThroughXSLT Property (Word)

Sets or returns a  **String** that specifies the path and file name for the Extensible Stylesheet Language Transformation (XSLT) to apply when a user saves a document.


## Syntax

 _expression_ . **XMLSaveThroughXSLT**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

The  **XMLSaveThroughXSLT** property is only applicable when the **[XMLUseXSLTWhenSaving](document-xmlusexsltwhensaving-property-word.md)** property is set to **True** . If the **XMLUseXSLTWhenSaving** property is set to **False** , Microsoft Word will ignore the **XMLSaveThroughXSLT** property.


## Example

The following example specifies that Word will use an XSLT when saving the active document, and then it specifies which XSLT to use.


```vb
ActiveDocument.XMLUseXSLTWhenSaving = True 
ActiveDocument.XMLSaveThroughXSLT = "c:\schemas\book.xsl"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

