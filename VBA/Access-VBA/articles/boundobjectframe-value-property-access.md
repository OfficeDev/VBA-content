---
title: BoundObjectFrame.Value Property (Access)
keywords: vbaac10.chm10906
f1_keywords:
- vbaac10.chm10906
ms.prod: access
api_name:
- Access.BoundObjectFrame.Value
ms.assetid: edafe10b-c207-527f-55a0-f71066fd9a85
ms.date: 06/08/2017
---


# BoundObjectFrame.Value Property (Access)

Gets of sets the value of the field that the control is bound to. Read/write  **Variant**.


## Syntax

 _expression_. **Value**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **Value** property for a bound object frame or a bound chart control is set to the value of the field that the control is bound to. Since these fields normally contain OLE objects or chart objects, which are stored as binary data, this value is usually meaningless.

The  **Value** property returns or sets a control's default property, which is the property that is assumed when you don't explicitly specify a property name.


 **Note**   The **Value** property is not the same as the **DefaultValue** property, which specifies the value that a property is assigned when a new record is created.


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

