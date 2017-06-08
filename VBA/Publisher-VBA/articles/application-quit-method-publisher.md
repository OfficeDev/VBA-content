---
title: Application.Quit Method (Publisher)
keywords: vbapb10.chm131129
f1_keywords:
- vbapb10.chm131129
ms.prod: publisher
api_name:
- Publisher.Application.Quit
ms.assetid: db5a02ec-e553-6de1-0e2c-4a9a512e68fe
ms.date: 06/08/2017
---


# Application.Quit Method (Publisher)

Quits Microsoft Publisher. This is equivalent to clicking  **Exit** on the **File** menu.


## Syntax

 _expression_. **Quit**

 _expression_A variable that represents an  **Application** object.


## Remarks

To avoid losing unsaved changes, use either the  **[Save](document-save-method-publisher.md)** or **[SaveAs](document-saveas-method-publisher.md)** method to save any open publication before calling the **Quit** method.


## Example

This example saves the open publication if there is one and then closes Publisher.


```vb
If Not (ActiveDocument Is Nothing) 
 ActiveDocument.Save 
End If 
Application.Quit
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

