---
title: JournalItem.ShowCategoriesDialog Method (Outlook)
keywords: vbaol11.chm1282
f1_keywords:
- vbaol11.chm1282
ms.prod: outlook
api_name:
- Outlook.JournalItem.ShowCategoriesDialog
ms.assetid: 3159ed4c-b272-764d-3ba7-ec5e7f8cd03e
ms.date: 06/08/2017
---


# JournalItem.ShowCategoriesDialog Method (Outlook)

Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.


## Syntax

 _expression_ . **ShowCategoriesDialog**

 _expression_ A variable that represents a **JournalItem** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new journal item, displays the item on the screen, and opens up the  **Show Categories** dialog box.


```vb
Sub JournalItem() 
 
 'Creates a journal item to access ShowCategoriesDialog 
 
 Dim olmyJournalItem As Outlook.JournalItem 
 
 'Create journal item 
 
 Set olmyJournalItem = Application.CreateItem(olJournalItem) 
 
 
 
 olmyJournalItem.Body = "Sales figure notes." 
 
 olmyJournalItem.Subject = "Sales Reports" 
 
 'Display the item 
 
 olmyJournalItem.Display 
 
 'Display the Show categories dialog 
 
 olmyJournalItem.ShowCategoriesDialog 
 
End Sub
```


## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)

