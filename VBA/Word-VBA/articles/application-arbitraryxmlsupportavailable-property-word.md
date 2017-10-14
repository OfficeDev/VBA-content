---
title: Application.ArbitraryXMLSupportAvailable Property (Word)
keywords: vbawd10.chm158335441
f1_keywords:
- vbawd10.chm158335441
ms.prod: word
api_name:
- Word.Application.ArbitraryXMLSupportAvailable
ms.assetid: 5cf53ae7-200b-811e-7946-4fefe825eaec
ms.date: 06/08/2017
---


# Application.ArbitraryXMLSupportAvailable Property (Word)

Returns a  **Boolean** that represents whether Microsoft Word accepts custom XML schemas. **True** indicates that Word accepts custom XML schemas.


## Syntax

 _expression_ . **ArbitraryXMLSupportAvailable**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

Microsoft Office Standard Edition 2003 includes XML support using the Word XML schema, but it does not provide support for custom XML schemas. Support for custom XML schemas is available only in the stand-alone release of Office Word 2003 or greater and in Office Professional Edition 2003 or greater. Use the  **ArbitraryXMLSupportAvailable** property to determine which release is installed.


## Example

The following code displays a message if the installed version of Word does not support custom XML schemas.


```
If Application.ArbitraryXMLSupportAvailable = False Then 
 MsgBox "Custom XML schemas are not " &; _ 
 "supported in this version of Microsoft Word."
```


## See also


#### Concepts


[Application Object](application-object-word.md)

