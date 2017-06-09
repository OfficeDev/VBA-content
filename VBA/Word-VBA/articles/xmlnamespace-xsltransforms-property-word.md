---
title: XMLNamespace.XSLTransforms Property (Word)
keywords: vbawd10.chm2293765
f1_keywords:
- vbawd10.chm2293765
ms.prod: word
api_name:
- Word.XMLNamespace.XSLTransforms
ms.assetid: 854900ad-74cc-ee1f-5e64-c42a2a439698
ms.date: 06/08/2017
---


# XMLNamespace.XSLTransforms Property (Word)

Returns an XSLTransforms collection that represents the Extensible Stylesheet Language Transformation (XSLT) files specified for use with a schema.


## Syntax

 _expression_ . **XSLTransforms**

 _expression_ An expression that returns an **[XMLNamespace](xmlnamespace-object-word.md)** object.


## Example

The following example adds the simplesample.xsl transform to the transforms for the SimpleSample schema.


 **Note**  The SimpleSample schema is included in the Smart Document Software Development Kit (SDK); however, there is no corresponding simplesample.xsl file. This code was created for example purposes only. For more information, refer to the Smart Document SDK on the Microsoft Developer Network (MSDN) Web site.


```vb
Dim objSchema As XMLNamespace 
Dim objTransform As XSLTransform 
 
Set objSchema = Application.XMLNamespaces("SimpleSample") 
Set objTransform = objSchema.XSLTransforms _ 
 .Add("c:\schemas\simplesample.xsl")
```


## See also


#### Concepts


[XMLNamespace Object](xmlnamespace-object-word.md)

