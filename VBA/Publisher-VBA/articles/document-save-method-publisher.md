---
title: Document.Save Method (Publisher)
keywords: vbapb10.chm196695
f1_keywords:
- vbapb10.chm196695
ms.prod: publisher
api_name:
- Publisher.Document.Save
ms.assetid: 89eae461-d1c2-b3ca-58b7-9528df8801d8
ms.date: 06/08/2017
---


# Document.Save Method (Publisher)

Saves the specified publication.


## Syntax

 _expression_. **Save**

 _expression_A variable that represents a  **Document** object.


## Remarks

If the publication has not been previously saved, calling the  **Save** method is equivalent to calling the **[SaveAs](document-saveas-method-publisher.md)** method with the **_FileName_** argument set to the value of the publication's **[Name](application-name-property-publisher.md)** property. If the publication has been previously saved, the **Save** method saves the current version of the publication in the format in which it was opened and in the location to which it was last saved.

Calling the  **Save** method always performs saves in the foreground regardless of whether background saves are enabled.


## Example

This example saves the active publication if it has changed since it was last saved.


```vb
If ActiveDocument.Saved = False Then ActiveDocument.Save
```


