---
title: ItemProperty Object (Outlook)
keywords: vbaol11.chm517
f1_keywords:
- vbaol11.chm517
ms.prod: outlook
api_name:
- Outlook.ItemProperty
ms.assetid: 3570d1f9-40ed-0a99-f63c-141134418c3b
ms.date: 06/08/2017
---


# ItemProperty Object (Outlook)

Represents information about a given item property for a Microsoft Outlook item object.


## Remarks

 Each item property defines a certain attribute of the item, such as the name, type, or value of the item. The **ItemProperty** object is a member of the **[ItemProperties](itemproperties-object-outlook.md)** collection.

Use  **ItemProperties.Item** ( _index_ ), where _index_ is the object's numeric position within the collection or it's name to return a single **ItemProperty** object.


## Example

The following example creates a reference to the first  **ItemProperty** object in the **ItemProperties** collection.


```
Sub NewMail() 
 
 'Creates a new MailItem and references the ItemProperties collection. 
 
 Dim objMail As MailItem 
 
 Dim objitems As ItemProperties 
 
 Dim objitem As ItemProperty 
 
 
 
 'Create a new mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 'Create a reference to the ItemProperties collection 
 
 Set objitems = objMail.ItemProperties 
 
 'Create reference to the first object in the collection 
 
 Set objitem = objitems.item(0) 
 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/4aff7ec9-14df-2ff3-7fd4-a8ab1ddac4ca%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/ceb37756-d7e4-fd27-372b-996669b8afa9%28Office.15%29.aspx)|
|[IsUserProperty](http://msdn.microsoft.com/library/6787380b-fe85-22d9-b95b-2b356bf84a21%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/f436386d-aa03-ab38-8ae1-1df0087f7495%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/7be24e63-3e5f-4ed9-a668-380077351636%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/f33cfcd0-f86b-d0cd-7d35-a21644bc5c42%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/12129828-ad07-08b9-9b32-d8b19aba7b6e%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/81144bd5-15d5-a233-6001-f8c80392850f%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[ItemProperty Object Members](http://msdn.microsoft.com/library/0de85516-c8e3-b985-0b7f-3098a0da7f2c%28Office.15%29.aspx)
