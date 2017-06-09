---
title: Control.Form Property (Access)
keywords: vbaac10.chm10139
f1_keywords:
- vbaac10.chm10139
ms.prod: access
api_name:
- Access.Control.Form
ms.assetid: 86612c78-65f8-dc56-77da-d031502822f7
ms.date: 06/08/2017
---


# Control.Form Property (Access)

You can use the  **Form** property to refer to a form or to refer to the form associated with a subformcontrol. Read-only **Form**.


## Syntax

 _expression_. **Form**

 _expression_ A variable that represents a **Control** object.


## Remarks

This property refers to a form object. It is read-only in all views.

This property is typically used to refer to the form or report contained in a subform control. For example, the following code uses the  **Form** property to access the OrderID control on a subform contained in the OrderDetails subform control.




```vb
Dim intOrderID As Integer 
intOrderID = Forms!Orders!OrderDetails.Form!OrderID
```

The next example calls a function from a property sheet by using the  **Form** property to refer to the active form that contains the control named CustomerID.




```
=MyFunction(Form!CustomerID)
```

When you use the  **Form** property in this manner, you are referring to the active form, and the name of the form isn't necessary.

The next example is the Visual Basic equivalent of the preceding example.




```vb
X = MyFunction(Forms!Customers!CustomerID)
```


 **Note**   When you use the **[Forms](forms-object-access.md)** collection, you must specify the name of the form.


## Example

The following example uses the  **Form** property to refer to a control on a subform.


```vb
Dim curTotalAmount As Currency 
 
curTotalAmount = Forms!Orders!OrderDetails.Form!TotalAmount 

```


## See also


#### Concepts


[Control Object](control-object-access.md)

