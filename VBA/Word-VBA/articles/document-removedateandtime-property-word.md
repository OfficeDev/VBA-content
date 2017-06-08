---
title: Document.RemoveDateAndTime Property (Word)
keywords: vbawd10.chm158007780
f1_keywords:
- vbawd10.chm158007780
ms.prod: word
api_name:
- Word.Document.RemoveDateAndTime
ms.assetid: 43520dad-0374-06c9-184e-da71de304360
ms.date: 06/08/2017
---


# Document.RemoveDateAndTime Property (Word)

Sets or returns a  **Boolean** indicating whether a document stores the date and time metadata for tracked changes. .


## Syntax

 _expression_ . **RemoveDateAndTime**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

 **True** removes date and time stamp information from tracked changes. **False** does not remove date and time stamp information from tracked changes. Use the **RemoveDateAndTime** property in conjunction with the **[RemovePersonalInformation](document-removepersonalinformation-property-word.md)** property to help remove personal information from the document properties.


## Example

The following example removes personal information from the active document, and it removes date and time information from any tracked changes in the document.


```vb
ActiveDocument.RemovePersonalInformation = True 
ActiveDocument.RemoveDateAndTime = True
```


## See also


#### Concepts


[Document Object](document-object-word.md)

