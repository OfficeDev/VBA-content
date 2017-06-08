---
title: Form.DefaultView Property (Access)
keywords: vbaac10.chm13352
f1_keywords:
- vbaac10.chm13352
ms.prod: access
api_name:
- Access.Form.DefaultView
ms.assetid: bb44eca9-1576-794a-0558-f67e2d37559b
ms.date: 06/08/2017
---


# Form.DefaultView Property (Access)

You can use the  **DefaultView** property to specify the opening view of a form. Read/write **Byte**.


## Syntax

 _expression_. **DefaultView**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **DefaultView** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Single Form|0|(Default) Displays one record at a time.|
|Continuous Forms|1|Displays multiple records (as many as will fit in the current window), each in its own copy of the form's detail section.|
|Datasheet|2|Displays the form fields arranged in rows and columns like a spreadsheet.|
|PivotTable|3|Displays the form as a PivotTable.|
|PivotChart|4|Displays the form as a PivotChart.|
|Split Form|5|Displayes a split view of a Single Form and a datasheet containing the form's records.|
The views displayed in the  **View** button list depend on the setting of the **ViewsAllowed** property. For example, if the **ViewsAllowed** property is set to Datasheet, Form View is disabled in the View button list and on the View menu.

The combination of these properties creates the following conditions.



|**DefaultView**|**ViewsAllowed**|**Description**|
|:-----|:-----|:-----|
|Single, Continuous Forms, or Datasheet|Both|Users can switch between Form view and Datasheet view.|
|Single or Continuous Forms|Form|Users can't switch from Form view to Datasheet view.|
|Single or Continuous Forms|Datasheet|Users can switch from Form view to Datasheet view but not back again.|
|Datasheet|Form|Users can switch from Datasheet view to Form view but not back again.|
|Datasheet|Datasheet|Users can't switch from Datasheet view to Form view.|

## See also


#### Concepts


[Form Object](form-object-access.md)

