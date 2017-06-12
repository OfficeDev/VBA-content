---
title: MailItem.BeforeCheckNames Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeCheckNames
ms.assetid: fac2b9c3-e662-d2d7-7b30-cd912b9ca891
ms.date: 06/08/2017
---


# MailItem.BeforeCheckNames Event (Outlook)

Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **BeforeCheckNames**( **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the name resolution process is not completed.|

## Remarks

You use the  **BeforeCheckNames** event in VBScript, but the event does not fire when an e-mail name is resolved on the form.

The event does not fire under the following circumstances:


- You customized a Journal Entry form and then resolved a contact in the  **Contacts** field.
    
- You customized a Contact form and then resolved a contact in the  **Contacts** field.
    
- You customized any type of form and Outlook automatically resolved the name in the background.
    
- You programmatically created and resolved a recipient.
    



## Example

This Visual Basic for Applications (VBA) example asks the user if the user wants to resolve names and returns  **False** to cancel the operation if the user answers no. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `SendMail()` procedure should be called before the event procedure can be called by Outlook.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Private Sub myItem_BeforeCheckNames(Cancel As Boolean) 
 
 If MsgBox("Do you want to resolve names now?", 4) = vbOK Then 
 
 Cancel = True 
 
 End If 
 
End Sub 
 
 
 
Public Sub SendMail() 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 myItem.Recipients.Add ("Dan Wilson") 
 
 myItem.Recipients.Add ("Nate Sun") 
 
 myItem.Body = "Good morning!" 
 
 myItem.Send 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

