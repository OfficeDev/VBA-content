---
title: Form.ServerFilterByForm Property (Access)
keywords: vbaac10.chm13483
f1_keywords:
- vbaac10.chm13483
ms.prod: access
api_name:
- Access.Form.ServerFilterByForm
ms.assetid: f9f8f28e-b67e-1f4e-a70b-c66169fca250
ms.date: 06/08/2017
---


# Form.ServerFilterByForm Property (Access)

You can use the  **ServerFilterByForm** property to specify or determine whether a form is opened in the Server Filter By Form window. Read/write **Boolean**.


## Syntax

 _expression_. **ServerFilterByForm**

 _expression_ A variable that represents a **Form** object.


## Remarks

The default value is  **False**.

You can remove a filter by using Visual Basic to set the  **ServerFilterByForm** property to **False**.




 **Note**  The  **ServerFilterByForm** property setting is ignored if the form's record source is a stored procedure.


## Example

The following example enables the "Order Lookup" form to be opened in a Microsoft Access Data Project in the Server Filter By Form window.


```vb
Forms("Order Lookup").ServerFilterByForm = True
```


## See also


#### Concepts


[Form Object](form-object-access.md)

