---
title: XMLSchemaReferences.IgnoreMixedContent Property (Word)
keywords: vbawd10.chm116129798
f1_keywords:
- vbawd10.chm116129798
ms.prod: word
api_name:
- Word.XMLSchemaReferences.IgnoreMixedContent
ms.assetid: 51515e72-a33c-b6c4-ee48-8252631dd438
ms.date: 06/08/2017
---


# XMLSchemaReferences.IgnoreMixedContent Property (Word)

Returns a  **Boolean** that represents whether Microsoft Word preforms validation on text nodes that have element siblings and specifies whether these text nodes are saved in XML when the **[XMLSaveDataOnly](http://msdn.microsoft.com/library/f670eb62-fa5a-a3ba-4db8-d4064014e936%28Office.15%29.aspx)** property is **True** . Read/write.


## Syntax

 _expression_ . **IgnoreMixedContent**

 _expression_ An expression that returns an **[XMLSchemaReferences](xmlschemareferences-object-word.md)** collection.


## Remarks

 **True** causes Word to ignore schema violations caused by text nodes that have element siblings; it also prevents these text nodes from being saved in XML when the **XMLSaveDataOnly** property is **True** , which helps to prevent text that was inserted by an Extensible Stylesheet Language Transformation (XLST) from being saved as part of the data. **False** raises validation errors on text nodes with element siblings.


## Example

The following example disables validation of XML and prevents text nodes that have elements as siblings from being saved as XML for the active document.


```vb
ActiveDocument.XMLSchemaReferences _ 
 .IgnoreMixedContent = True
```


## See also


#### Concepts


[XMLSchemaReferences Collection](xmlschemareferences-object-word.md)

