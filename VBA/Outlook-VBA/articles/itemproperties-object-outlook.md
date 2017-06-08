---
title: ItemProperties Object (Outlook)
keywords: vbaol11.chm530
f1_keywords:
- vbaol11.chm530
ms.prod: outlook
api_name:
- Outlook.ItemProperties
ms.assetid: 34a110ed-6617-72da-1e98-a9773c705b40
ms.date: 06/08/2017
---


# ItemProperties Object (Outlook)

A collection of all properties associated with the item.


## Remarks

Use the  **[ItemProperties](http://msdn.microsoft.com/library/620e3af5-0c11-bd78-a98f-b08b36857113%28Office.15%29.aspx)** property to return the **ItemProperties** collection. Use **ItemProperties.Item** ( _index_ ), where _index_ is the name of the object or the numeric position of the item within the collection, to return a single **[ItemProperty](itemproperty-object-outlook.md)** object.


 **Note**  The  **ItemProperties** collection is zero-based, meaning that the first item in the collection is referenced by 0.

Use the  **[Add](http://msdn.microsoft.com/library/317daeba-e34c-8458-2492-c434707fa805%28Office.15%29.aspx)** method to add a new item property to the **ItemProperties** collection. Use the **[Remove](http://msdn.microsoft.com/library/51d0320b-99f4-60df-4646-b8e365813d2f%28Office.15%29.aspx)** method to remove an item property from the **ItemProperties** collection.


 **Note**   You can only add or remove custom properties. Custom properties are denoted by the **[IsUserProperty](http://msdn.microsoft.com/library/6787380b-fe85-22d9-b95b-2b356bf84a21%28Office.15%29.aspx)**.


## Example

The following example creates a new  **[MailItem](http://msdn.microsoft.com/library/14197346-05d2-0250-fa4c-4a6b07daf25f%28Office.15%29.aspx)** object and stores its **ItemProperties** collection in a variable called `objItems`.


```
Sub ItemProperty() 
 
 'Creates a new MailItem and access its properties 
 
 Dim objMail As MailItem 
 
 Dim objItems As ItemProperties 
 
 Dim objItem As ItemProperty 
 
 
 
 'Create the mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 'Create a reference to the item properties collection 
 
 Set objItems = objMail.ItemProperties 
 
 'Create a reference to the item property page 
 
 Set objItem = objItems.item(0) 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/317daeba-e34c-8458-2492-c434707fa805%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/51bb7900-d3fc-650d-d43b-0da14e13ca5a%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/51d0320b-99f4-60df-4646-b8e365813d2f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/b5e8e499-136c-a41e-cfe8-73637b44b8b2%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/356d4e84-9e5c-10fc-bced-f7f176378bd9%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/4838ad3a-a06e-b7e2-0566-734c9b79515c%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/2756ca03-4ba8-583c-12a5-1cff103417eb%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/5dde3402-b791-e0f7-e4fe-10bb85e5284a%28Office.15%29.aspx)|

## See also


#### Other resources


[ItemProperties Object Members](http://msdn.microsoft.com/library/9c18dfa4-b0df-0a01-cac8-cb4ef7a4f2b5%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
