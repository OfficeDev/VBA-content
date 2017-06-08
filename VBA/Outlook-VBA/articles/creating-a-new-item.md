---
title: Creating a New Item
keywords: olfm10.chm3077112
f1_keywords:
- olfm10.chm3077112
ms.prod: outlook
ms.assetid: 3e7e5c7d-d0f8-36f4-c126-9f4ef73057a3
ms.date: 06/08/2017
---


# Creating a New Item

To create a new item, use the  **[CreateItem](application-createitem-method-outlook.md)** method of the **[Application](application-object-outlook.md)** object. This method returns an object that you can then use to work with the item.

The following Microsoft Visual Basic for Applications example shows how to create a mail message, add text to its subject and body, and display it. To use this sample, create a command button named Command1 on a form.



```vb
Private Sub Command1_Click() 
 Dim myOLItem As Outlook.MailItem 
 
 Set myOLItem = Application.CreateItem(olMailItem) 
 With myOLItem 
 .Subject = "Sample item" 
 .Body = "This is a sample message." 
 End With 
 myOLItem.Display 
End Sub
```

The following example shows how to perform the same task using VBScript in a form.



```vb
Sub CommandButton1_Click() 
 Set myOLItem = Application.CreateItem(0) 
 myOLItem.Subject = "Sample item" 
 myOLItem.Body = "This is a sample message." 
 myOLItem.Display 
End Sub
```


