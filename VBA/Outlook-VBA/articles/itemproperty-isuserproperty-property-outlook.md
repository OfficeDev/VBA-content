---
title: ItemProperty.IsUserProperty Property (Outlook)
keywords: vbaol11.chm529
f1_keywords:
- vbaol11.chm529
ms.prod: outlook
api_name:
- Outlook.ItemProperty.IsUserProperty
ms.assetid: 6787380b-fe85-22d9-b95b-2b356bf84a21
ms.date: 06/08/2017
---


# ItemProperty.IsUserProperty Property (Outlook)

Returns a  **Boolean** value that indicates if the item property is a custom property created by the user. Read-only.


## Syntax

 _expression_ . **IsUserProperty**

 _expression_ A variable that represents an **ItemProperty** object.


## Example

The following example displays the names of all properties created by the user. The subroutine  `DisplayUserProps` accepts an **[ItemProperties](itemproperties-object-outlook.md)** collection and searches through it, displaying the names of all **[ItemProperty](itemproperty-object-outlook.md)** objects where the **IsUserProperty** value is **True** . The **ItemProperties** collection is zero-based. In other words, the first object in the collection is accessed with an index value of zero (0).


```vb
Sub ItemProperty() 
 'Creates a new mail item and access it's properties 
 Dim objMail As MailItem 
 Dim objitems As ItemProperties 
 
 'Create the mail item 
 Set objMail = Application.CreateItem(olMailItem) 
 'Create a reference to the item properties collection 
 Set objitems = objMail.ItemProperties 
 'Create a reference to the item property page 
 Call DisplayUserProps(objitems) 
End Sub 
 
Sub DisplayUserProps(ByVal objitems As ItemProperties) 
 'Displays the names of all user-created item properties in the collection 
 For i = 0 To objitems.Count - 1 
 'Display name of property if it was created by the user 
 If objitems.Item(i).IsUserProperty = True Then 
 MsgBox "The property " &; objitems(i).Name &; " was created by the user." 
 End If 
 Next i 
End Sub
```


## See also


#### Concepts


[ItemProperty Object](itemproperty-object-outlook.md)

