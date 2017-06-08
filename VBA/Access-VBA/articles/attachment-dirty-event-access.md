---
title: Attachment.Dirty Event (Access)
keywords: vbaac10.chm14023
f1_keywords:
- vbaac10.chm14023
ms.prod: access
api_name:
- Access.Attachment.Dirty
ms.assetid: d211238b-cbe4-f0ef-471b-33c1ced1aa9b
ms.date: 06/08/2017
---


# Attachment.Dirty Event (Access)

The  **Dirty** event occurs when the content of the specified control changes.


## Syntax

 _expression_. **Dirty**( ** _Cancel_**, )

 _expression_ A variable that represents an **Attachment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the  **Dirty** event occurs. Setting the _Cancel_ argument to **True** (?1) cancels the **Dirty** event.|

## Remarks

Examples of this event include entering a character directly in the text box or combo box or changing the control's  **Text** property setting by using a macro or Visual Basic.


- Modifying a record within a form by using a macro or Visual Basic doesn't trigger this event. You must type the data directly into the record or set the control's  **Text** property.
    
- This event applies only to bound forms, not an unbound form or report.
    
To run a macro or event procedure when this event occurs, set the  **OnDirty** property to the name of the macro or to [Event Procedure].

Canceling the  **Dirty** event will cause the changes to the current record to be rolled back. It is equivalent to pressing the ESC key.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

