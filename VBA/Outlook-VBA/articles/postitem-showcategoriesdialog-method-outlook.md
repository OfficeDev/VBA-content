---
title: PostItem.ShowCategoriesDialog Method (Outlook)
keywords: vbaol11.chm1560
f1_keywords:
- vbaol11.chm1560
ms.prod: outlook
api_name:
- Outlook.PostItem.ShowCategoriesDialog
ms.assetid: 00483040-7c23-e920-3d97-1ac456c25b05
ms.date: 06/08/2017
---


# PostItem.ShowCategoriesDialog Method (Outlook)

Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.


## Syntax

 _expression_ . **ShowCategoriesDialog**

 _expression_ A variable that represents a **PostItem** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new post item, displays the item on the screen, and opens up the  **Show Categories** dialog box.


```vb
Sub PostItem() 
 
 'Creates a post item to access ShowCategoriesDialog 
 
 Dim olmyPostItem As Outlook.PostItem 
 
 'Create post item 
 
 Set olmyPostItem = Application.CreateItem(olPostItem) 
 
 
 
 olmyPostItem.Body = "Please comment on these sales figures." 
 
 olmyPostItem.Subject = "Sales Reports" 
 
 'Display the item 
 
 olmyPostItem.Display 
 
 'Display the Show categories dialog 
 
 olmyPostItem.ShowCategoriesDialog 
 
End Sub
```


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

