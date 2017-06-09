---
title: DocumentProperties.Parent Property (Office)
keywords: vbaof11.chm250011
f1_keywords:
- vbaof11.chm250011
ms.prod: office
api_name:
- Office.DocumentProperties.Parent
ms.assetid: e1239ffa-b89e-e78f-4009-d576c473d477
ms.date: 06/08/2017
---


# DocumentProperties.Parent Property (Office)

Gets the  **Parent** object for the **DocumentProperties** object. Read-only.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **DocumentProperties** object.


### Return Value

Object


## Example

This example displays the name of the parent object for a document property. You must pass a valid  **DocumentProperty** object to the procedure.


```
Sub DisplayParent(dp as DocumentProperty) 
 MsgBox dp.Parent.Name 
End Sub
```


## See also


#### Concepts


[DocumentProperties Object](documentproperties-object-office.md)
#### Other resources


[DocumentProperties Object Members](documentproperties-members-office.md)

