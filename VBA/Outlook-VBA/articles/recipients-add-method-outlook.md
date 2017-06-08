---
title: Recipients.Add Method (Outlook)
keywords: vbaol11.chm232
f1_keywords:
- vbaol11.chm232
ms.prod: outlook
api_name:
- Outlook.Recipients.Add
ms.assetid: 7c285291-0f92-ca8d-1c7b-a71ace83ac84
ms.date: 06/08/2017
---


# Recipients.Add Method (Outlook)

Creates a new recipient in the  **[Recipients](recipients-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**( **_Name_** )

 _expression_ A variable that represents a **Recipients** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the recipient; it can be a string representing the display name, the alias, or the full SMTP e-mail address of the recipient.|

### Return Value

A  **[Recipient](recipient-object-outlook.md)** object that represents the new recipient.


## Example

This VBA example creates a new mail message, uses the Add method to add 'Dan Wilson' as a To recipient, and displays the message. To run this example without errors, replace 'Dan Wilson' with a valid recipient name.


```vb
Sub CreateStatusReportToBoss() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myRecipient As Outlook.Recipient 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myRecipient = myItem.Recipients.Add("Dan Wilson") 
 
 myItem.Subject = "Status Report" 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[Recipients Object](recipients-object-outlook.md)

