---
title: MailItem.ShowCategoriesDialog Method (Outlook)
keywords: vbaol11.chm1374
f1_keywords:
- vbaol11.chm1374
ms.prod: outlook
api_name:
- Outlook.MailItem.ShowCategoriesDialog
ms.assetid: 212dfd98-c0a2-7f94-249f-ba9baec34882
ms.date: 06/08/2017
---


# MailItem.ShowCategoriesDialog Method (Outlook)

Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.


## Syntax

 _expression_ . **ShowCategoriesDialog**

 _expression_ A variable that represents a **MailItem** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new mail item, displays the item on the screen, and opens up the  **Show Categories** dialog box.


```vb
Sub MailItem() 
 
 'Creates a mail item to access ShowCategoriesDialog 
 
 Dim olmyMailItem As Outlook.MailItem 
 
 'Create mail item 
 
 Set olmyMailItem = Application.CreateItem(olMailItem) 
 
 
 
 olmyMailItem.Body = "Can you help me with these sales figures." 
 
 olmyMailItem.Recipients.Add ("Jeff Smith") 
 
 olmyMailItem.Subject = "Sales Reports" 
 
 'Display the item 
 
 olmyMailItem.Display 
 
 'Display the Show categories dialog 
 
 olmyMailItem.ShowCategoriesDialog 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

