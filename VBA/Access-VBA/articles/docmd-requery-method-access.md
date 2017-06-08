---
title: DoCmd.Requery Method (Access)
keywords: vbaac10.chm4170
f1_keywords:
- vbaac10.chm4170
ms.prod: access
api_name:
- Access.DoCmd.Requery
ms.assetid: 6869c39f-b43f-ad83-4140-67b763342605
ms.date: 06/08/2017
---


# DoCmd.Requery Method (Access)

Carries out the Requery action in Visual Basic.


## Syntax

 _expression_. **Requery**( ** _ControlName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ControlName_|Optional|**Variant**|A string expression that's the name of a control on the active object.|

## Remarks

You can use the Requery action to update the data in a specified control on the active object by requerying the source of the control. If no control is specified, this action requeries the source of the object itself. Use this action to ensure that the active object or one of its controls displays the most current data.

If you leave the Control Name argument blank, the Requery action has the same effect as pressing SHIFT+F9 when the object has the focus. If a subform control has the focus, this action requeries only the source of the subform (just as pressing SHIFT+F9 does).

If you want to requery a control that isn't on the active object, you must use the  **Requery** method in Visual Basic, not the Requery action or its corresponding **Requery** method of the **DoCmd** object. The **Requery** method in Visual Basic is faster than the Requery action or the **DoCmd.Requery** method. In addition, when you use the Requery action or the **DoCmd.Requery** method, Microsoft Access closes the query and reloads it from the database, but when you use the **Requery** method, Access reruns the query without closing and reloading it. Note that the ActiveX Data Object (ADO) **Requery** method works the same way as the Access **Requery** method.


## Example

The following example uses the  **Requery** method to update the EmployeeList control:


```vb
DoCmd.Requery "EmployeeList"
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

