---
title: XMLNamespace.DefaultTransform Property (Word)
keywords: vbawd10.chm2293766
f1_keywords:
- vbawd10.chm2293766
ms.prod: word
api_name:
- Word.XMLNamespace.DefaultTransform
ms.assetid: a43c9869-98f0-0a18-8e3c-eb4930553367
ms.date: 06/08/2017
---


# XMLNamespace.DefaultTransform Property (Word)

 Returns an **[XSLTransform](xsltransform-object-word.md)** object that represents the default Extensible Stylesheet Language Transformation (XSLT) file to use when opening a document from an XML schema for a particular namespace.


## Syntax

 _expression_ . **DefaultTransform**

 _expression_ An expression that returns an **[XMLNamespace](xmlnamespace-object-word.md)** object.


## Example

The following example returns the default XSLT for the first schema in the Schema Library that Microsoft Word will use to open XML files associated with that schema's namespace. This example assumes that the first schema has one or more applied XSLT files.


```vb
Dim objXSLT As XSLTransform 
 
Set objXSLT = Application.XMLNamespaces(1).DefaultTransform
```


## See also


#### Concepts


[XMLNamespace Object](xmlnamespace-object-word.md)

