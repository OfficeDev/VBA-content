---
title: MailItem.Close Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.Close
ms.assetid: 95caf7b5-d139-8b8b-bcd2-874243c4ed50
ms.date: 06/08/2017
---


# MailItem.Close Event (Outlook)

Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.


## Syntax

 _expression_ . **Close**( **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the close operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the close operation isn't completed and the inspector is left open.

If you use the  **[Close](mailitem-close-method-outlook.md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.


## Example

This Microsoft Visual Basic for Applications (VBA) example tests for the  **Close** event and if the item has not been **[Saved](mailitem-saved-property-outlook.md)** , it uses the **[Save](mailitem-save-method-outlook.md)** method to save the item without prompting the user.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Public Sub Initalize_Handler() 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
End Sub 
 
 
 
Private Sub myItem_Close(Cancel As Boolean) 
 
 If Not myItem.Saved Then 
 
 myItem.Save 
 
 MsgBox " The item was saved." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

