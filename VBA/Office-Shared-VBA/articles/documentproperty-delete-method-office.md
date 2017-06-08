---
title: DocumentProperty.Delete Method (Office)
keywords: vbaof11.chm250004
f1_keywords:
- vbaof11.chm250004
ms.prod: office
api_name:
- Office.DocumentProperty.Delete
ms.assetid: 2a9ac097-0156-007f-2b4b-62a34b240f71
ms.date: 06/08/2017
---


# DocumentProperty.Delete Method (Office)

Removes a custom document property.


## Syntax

 _expression_. **Delete**

 _expression_ Required. A variable that represents a **[DocumentProperty](documentproperty-object-office.md)** object.


## Remarks

You cannot delete a built-in document property.


## Example

This example deletes a custom document property. For this example to run properly, you must have a custom DocumentProperty object named "CustomNumber".


```
ActiveDocument.CustomDocumentProperties("CustomNumber").Delete
```


## See also


#### Concepts


[DocumentProperty Object](documentproperty-object-office.md)
#### Other resources


[DocumentProperty Object Members](documentproperty-members-office.md)

