---
title: CustomControl.OLEClass Property (Access)
keywords: vbaac10.chm12010
f1_keywords:
- vbaac10.chm12010
ms.prod: access
api_name:
- Access.CustomControl.OLEClass
ms.assetid: d9aad7b9-6388-3365-881a-6e42ebebcfd6
ms.date: 06/08/2017
---


# CustomControl.OLEClass Property (Access)

You can use the  **OLEClass** property to obtain a description of the kind of OLE object contained in a chart control or an unbound object frame. Read-only **String**.


## Syntax

 _expression_. **OLEClass**

 _expression_ A variable that represents a **CustomControl** object.


## Remarks

This property is set automatically in the control's property sheet to a string expression when you click  **Object** on the **Insert** menu to add an OLE object to a form. The **OLEClass** property setting is read-only in all views.


 **Note**  If you are using Automation (formerly called OLE Automation) and need to specify a name to refer to the OLE object, use the  **Class** property.

The  **OLEClass** property and the **Class** property are similar but not identical. The **OLEClass** property setting is a general description of the OLE object whereas the **Class** property setting is the name used to refer to the OLE object in Visual Basic. Examples of **OLEClass** property settings are Microsoft Excel Chart, Microsoft Word Document, and Paintbrush Picture.


## Example

The following example displays a message indicating the OLE class for the "Customer Picture" unbound object frame on the "Order Entry" form.


```vb
MsgBox "The OLE class = " &; Forms("Order Entry").Controls("Customer Picture").OLEClass
```


## See also


#### Concepts


[CustomControl Object](customcontrol-object-access.md)

