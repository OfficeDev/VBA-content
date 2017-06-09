---
title: XMLSchemaReferences.Validate Method (Word)
keywords: vbawd10.chm116129892
f1_keywords:
- vbawd10.chm116129892
ms.prod: word
api_name:
- Word.XMLSchemaReferences.Validate
ms.assetid: 66e4ea2d-e26c-be4c-fe1d-d240449f30f3
ms.date: 06/08/2017
---


# XMLSchemaReferences.Validate Method (Word)

Validates all the XML schemas that are attached to a document.


## Syntax

 _expression_ . **Validate**

 _expression_ An expression that returns an **[XMLSchemaReferences](xmlschemareferences-object-word.md)** object.


### Return Value

Nothing


## Remarks

When you run the  **Validate** method, Microsoft Word populates the **[XMLSchemaViolations](http://msdn.microsoft.com/library/9bed9233-4b6b-fe11-d681-8c9f72f99449%28Office.15%29.aspx)** property of the **[Document](document-object-word.md)** object with a collection of the XML nodes that have validation errors.


## See also


#### Concepts


[XMLSchemaReferences Collection](xmlschemareferences-object-word.md)

