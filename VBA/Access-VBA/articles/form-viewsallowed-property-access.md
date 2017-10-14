---
title: Form.ViewsAllowed Property (Access)
keywords: vbaac10.chm13353
f1_keywords:
- vbaac10.chm13353
ms.prod: access
api_name:
- Access.Form.ViewsAllowed
ms.assetid: 2aa001e0-ea0d-4ef3-f8d2-fdd301502c96
ms.date: 06/08/2017
---


# Form.ViewsAllowed Property (Access)

You can use the  **ViewsAllowed** property to specify whether users can switch between Datasheet view and Form view by clicking the **Form View** or **Datasheet View** command on the **View** menu or by clicking the arrow next to the **View** button and clicking **Form View** or **Datasheet View**. Read/write **Byte**.


## Syntax

 _expression_. **ViewsAllowed**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **ViewsAllowed** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|(Default) Users can switch between Form view and Datasheet view.|
|1|Users can't switch to Datasheet view from Form view.|
|2|Users can't switch to Form view from Datasheet view.|
The views displayed in the  **View** button list and on the **View** menu depend on the setting of the **ViewsAllowed** property. For example, if the **ViewsAllowed** property is set to Datasheet, **Form View** is disabled in the **View** button list and on the **View** menu.

The combination of these properties creates the following conditions.



|**DefaultView**|**ViewsAllowed**|**Description**|
|:-----|:-----|:-----|
|Single, Continuous Forms, or Datasheet|Both|Users can switch between Form view and Datasheet view.|
|Single or Continuous Forms|Form|Users can't switch from Form view to Datasheet view.|
|Single or Continuous Forms|Datasheet|Users can switch from Form view to Datasheet view but not back again.|
|Datasheet|Form|Users can switch from Datasheet view to Form view but not back again.|
|Datasheet|Datasheet|Users can't switch from Datasheet view to Form view.|

## Example

The following example prints a message in the Immediate window indicating the state of how users can switch between Datasheet view and Form view for the "Switchboard" form.


```vb
Debug.Print "ViewsAllowed = " &; Forms.Item("Switchboard").ViewsAllowed
```


## See also


#### Concepts


[Form Object](form-object-access.md)

