---
title: MailItem.Write Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.Write
ms.assetid: b4c5fc80-e197-8d82-ebb0-148675ea7cdd
ms.date: 06/08/2017
---


# MailItem.Write Event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](mailitem-save-method-outlook.md)** or **[SaveAs](mailitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

 _expression_ . **Write**( **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the save operation is not completed.


## Example

This Visual Basic for Applications (VBA) example uses the  **Write** event and warns the user that the item is about to be saved and will overwrite any existing item and, depending on the user's response, either allows the operation to continue or stops it. If this event is canceled, Microsoft Outlook displays an error message. Therefore, you need to capture this event in your code. One way to do this is shown below. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `Initialize_Handler()` subroutine must be called before the event procedure can be called by Outlook.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Private Sub myItem_Write(Cancel As Boolean) 
 
 Dim myResult As Integer 
 
 myItem = "The item is about to be saved. Do you wish to overwrite the existing item?" 
 
 myResult = MsgBox(myItem, vbYesNo, "Save") 
 
 If myResult = vbNo Then 
 
 Cancel = True 
 
 End If 
 
End Sub 
 
 
 
Public Sub Initalize_Handler() 
 
 Const strCancelEvent = "Application-defined or object-defined error" 
 
 
 
 On Error GoTo ErrHandler 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 myItem.Save 
 
 Exit Sub 
 
 
 
 ErrHandler: 
 
 MsgBox Err.Description 
 
 If Err.Description = strCancelEvent Then 
 
 MsgBox "The event was cancelled." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

