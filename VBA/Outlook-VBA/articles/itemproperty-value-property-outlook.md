---
title: ItemProperty.Value Property (Outlook)
keywords: vbaol11.chm527
f1_keywords:
- vbaol11.chm527
ms.prod: outlook
api_name:
- Outlook.ItemProperty.Value
ms.assetid: 81144bd5-15d5-a233-6001-f8c80392850f
ms.date: 06/08/2017
---


# ItemProperty.Value Property (Outlook)

Returns or sets a  **Variant** indicating the value for the specified custom or explicit built-in property. Read/write.


## Syntax

 _expression_ . **Value**

 _expression_ A variable that represents an **ItemProperty** object.


## Remarks

Even though  **ItemProperty.Value** allows you to get or set an explicit built-in property or a custom property, you can reference explicit built-in properties directly from the parent object, for example, `ContactItem.Body`. For more information on accessing properties in Outlook, see [Properties Overview](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx).


## Example

The following Visual Basic for Applications (VBA) example creates a contact item and sets its  **Body** property


```vb
Sub ValueItemProperty() 
 
 Dim cti As Outlook.ContactItem 
 
 Dim itms As Outlook.ItemProperties 
 
 Dim itm As Outlook.ItemProperty 
 
 
 
 Set cti = Application.CreateItem(olContactItem) 
 
 cti.FullName = "Dan Wilson" 
 
 Set itms = cti.ItemProperties 
 
 Set itm = itms.Item("Body") 
 
 itm.Value = "My friend from school" 
 
 cti.Save 
 
 cti.Display 
 
End Sub
```


## See also


#### Concepts


[ItemProperty Object](itemproperty-object-outlook.md)

