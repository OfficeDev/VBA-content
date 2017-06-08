---
title: MailItem.Forward Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.Forward
ms.assetid: 29426284-471b-95bb-be67-a3ca3f9a0d79
ms.date: 06/08/2017
---


# MailItem.Forward Event (Outlook)

Occurs when the user selects the  **Forward** action for an item, or when the **Forward** method is called for the item, which is an instance of the parent object.


## Syntax

 _expression_ . **Forward**( **_Forward_** , **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in Microsoft Visual Basic Scripting Edition (VBScript).)  **False** when the event occurs. If the event procedure sets this argument to **True** , the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False** , the forward action is not completed and the new item is not displayed.


## Example

This Microsoft Visual Basic for Applications (VBA) example uses the  **Forward** event to disable forwarding on an item that has the subject "Do not forward" by setting the Cancel argument to **True** and it also displays a message that the item may not be forwarded. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `Initialize_Handler()` routine should be called before the event procedure can be called by Microsoft Outlook. A e-mail item must be open when you run `Initialize_Handler()`.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Public Sub Initialize_Handler() 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
End Sub 
 
 
 
Private Sub myItem_Forward(ByVal Forward As Object, Cancel As Boolean) 
 
 If myItem.Subject = "Do not forward" Then 
 
 MsgBox "You may not forward this message!" 
 
 Cancel = True 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

