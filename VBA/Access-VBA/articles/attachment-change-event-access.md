---
title: Attachment.Change Event (Access)
keywords: vbaac10.chm14024
f1_keywords:
- vbaac10.chm14024
ms.prod: access
api_name:
- Access.Attachment.Change
ms.assetid: 5b34517d-f3a8-a10d-1bc3-ed3bc8ecc484
ms.date: 06/08/2017
---


# Attachment.Change Event (Access)

The  **Change** event occurs when the contents of the specified control change.


## Syntax

 _expression_. **Change**

 _expression_ A variable that represents an **Attachment** object.


### Return Value

nothing


## Remarks


 **Note**  

To run a macro or event procedure when this event occurs, set the  **OnChange** property to the name of the macro or to [Event Procedure].

By running a macro or event procedure when a Change event occurs, you can coordinate data display among controls. You can also display data or a formula in one control and the results in another control.

The Change event doesn't occur when a value changes in a calculated control.


 **Note**  A Change event can cause a cascading event. This occurs when a macro or event procedure that runs in response to the control's Change event alters the control's contents — for example, by changing a property setting that determines the control's value, such as the  **Text** property for a text box. To prevent a cascading event:


- If possible, avoid attaching a Change macro or event procedure to a control that alters the control's contents.
    
- Avoid creating two or more controls having Change events that affect each other — for example, two text boxes that update each other.
    

## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

