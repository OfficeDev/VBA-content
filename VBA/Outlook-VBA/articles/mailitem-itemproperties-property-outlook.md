---
title: MailItem.ItemProperties Property (Outlook)
keywords: vbaol11.chm1371
f1_keywords:
- vbaol11.chm1371
ms.prod: outlook
api_name:
- Outlook.MailItem.ItemProperties
ms.assetid: 620e3af5-0c11-bd78-a98f-b08b36857113
ms.date: 06/08/2017
---


# MailItem.ItemProperties Property (Outlook)

Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.


## Syntax

 _expression_ . **ItemProperties**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

The  **ItemProperties** collection is a zero-based collection, meaning that the first object in the collection is referenced by the index 0.


## Example

The following Microsoft Visual Basic for Applications (VBA) example returns the  **ItemProperties** collection associated with a **[MailItem](mailitem-object-outlook.md)** object.


```vb
Sub ItemProperty() 
 
 'Creates a new e-mail item and accesses its properties. 
 
 Dim objMail As Outlook.MailItem 
 
 Dim objItems As Outlook.ItemProperties 
 
 Dim objItem As Outlook.ItemProperty 
 
 
 
 'Create the e-mail item. 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 'Create a reference to the e-mail item's properties collection. 
 
 Set objItems = objMail.ItemProperties 
 
 'Create a reference to the third e-mail item property. 
 
 Set objItem = objItems.Item(2) 
 
 MsgBox objItem.Name &; " = " &; objItem.Value 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

