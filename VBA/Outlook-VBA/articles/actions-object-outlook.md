---
title: Actions Object (Outlook)
keywords: vbaol11.chm144
f1_keywords:
- vbaol11.chm144
ms.prod: outlook
api_name:
- Outlook.Actions
ms.assetid: b0903aa4-9b75-5311-d0a5-5ff4a5e29c79
ms.date: 06/08/2017
---


# Actions Object (Outlook)

Contains a collection of  **[Action](action-object-outlook.md)** objects that represent all the specialized actions that can be executed on an Outlook item.


## Remarks

Use the  **Actions** property of any Outlook item, such as **[MailItem](http://msdn.microsoft.com/library/14197346-05d2-0250-fa4c-4a6b07daf25f%28Office.15%29.aspx)**, to return the **Actions** object.

Use  **Actions** ( _index_ ), where _index_ is the name of an available action, to return a single **Action** object.


## Example

The following Visual Basic for Applications (VBA) example uses the Reply action of a particular item to send a reply.


```
myItem = CreateItem(olMailItem) 
 
Set myReply = myItem.Actions("Reply").Execute
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/aaf539c4-d60a-867f-086b-3cef7632a6f2%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/e4c10f5e-014f-46d5-e5a9-2a70c9399d5f%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/a44c382b-0eff-2033-da91-05bee0e210b2%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/823b9111-fb73-581b-18e0-68f34a71fa3e%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/fe55f517-bb09-5d57-0ca1-f50fe1d482c2%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/0ba24d51-b057-9960-18e0-cb88a5edcdd5%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/c92854dd-19f7-39d4-9b81-76645c032577%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/21792c3f-9669-2f68-7a47-bac172d16620%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[Actions Object Members](http://msdn.microsoft.com/library/f4791bd5-87bb-ac1e-0acc-709cf5f91e36%28Office.15%29.aspx)
