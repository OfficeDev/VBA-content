---
title: Form.SubdatasheetHeight Property (Access)
keywords: vbaac10.chm13510
f1_keywords:
- vbaac10.chm13510
ms.prod: access
api_name:
- Access.Form.SubdatasheetHeight
ms.assetid: 0db2e4b5-e64b-6f55-ebfa-bcce98734491
ms.date: 06/08/2017
---


# Form.SubdatasheetHeight Property (Access)

You can use the  **SubdatasheetHeight** property to specify or determine the default display height of a subdatasheet when expanded. Read/write **Integer**.


## Syntax

 _expression_. **SubdatasheetHeight**

 _expression_ A variable that represents a **Form** object.


## Remarks

the  **SubdatasheetHeight** property's value is expressed in twips.

To set the  **SubdatasheetHeight** property by using Visual Basic, you must first create the property by using the DAO **CreateProperty** method.

If the subdatasheet includes more records than the height setting can accommodate, a vertical scrollbar is displayed.

The  **SubdatasheetHeight** property setting includes the New Record row if adding new records is supported. It does not include the column header row or scrollbar region.

The  **SubdatasheetHeight** and **SubdatasheetExpanded** properties take effect on the subform control when the form is in datasheet view.


## Example

The following example resizes the height of the subdatasheet in the "Purchase Orders" form (containing a subform) to show only one line of the subdatasheet at a time (measured at about 400 twips), accompanied by a vertical scrollbar. The number 400 is arbitrary, and will vary based on monitor resolution and default font size. This behavior can only be seen in Datasheet View.


```vb
Forms("Purchase Orders").SubdatasheetHeight = 400
```


## See also


#### Concepts


[Form Object](form-object-access.md)

