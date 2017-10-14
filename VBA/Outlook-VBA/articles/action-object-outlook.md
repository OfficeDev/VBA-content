---
title: Action Object (Outlook)
keywords: vbaol11.chm9
f1_keywords:
- vbaol11.chm9
ms.prod: outlook
api_name:
- Outlook.Action
ms.assetid: 22bd8d4a-9cf4-bd37-011b-8da3dfadf761
ms.date: 06/08/2017
---


# Action Object (Outlook)

Represents a specialized action (for example, the voting options response) that can be executed on an Outlook item.


## Remarks

The  **Action** object is a member of the **[Actions](actions-object-outlook.md)** collection.

Use  **[Actions](http://msdn.microsoft.com/library/1b7bb1c0-334f-826a-fd6b-8fc3f2fe5d64%28Office.15%29.aspx)** ( _index_ ), where _index_ is the name of an available action, to return a single **Action** object from the **Actions** collection object of an Outlook item, such as **[MailItem](http://msdn.microsoft.com/library/14197346-05d2-0250-fa4c-4a6b07daf25f%28Office.15%29.aspx)**.


## Example

The following Visual Basic for Applications (VBA) example uses the Reply action of a particular item to send a reply.


```
myItem = CreateItem(olMailItem) 
 
Set myReply = myItem.Actions("Reply").Execute
```

The following Visual Basic for Applications example does the same thing, using a different reply style for the reply.




```
myItem = CreateItem(olMailItem) 
 
myItem.Actions("Reply").ReplyStyle = _ 
 
 olIncludeOriginalText 
 
Set myReply = myItem.Actions("Reply").Execute
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/96d498d2-9035-f31c-e2d1-3431e15f39db%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/29dd0c5c-ed5f-b2cc-45b0-1c8c348239bb%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/f89f7f23-1231-aa53-d720-6571145a807d%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/cdba7120-30d8-621f-415d-4c4b4101b4bc%28Office.15%29.aspx)|
|[CopyLike](http://msdn.microsoft.com/library/4cde4458-1bf1-7673-1c5f-d3d9c4e9b8f6%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/090b5fdf-42fb-4da8-fb8f-74accaf1dc80%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/a1a1eaeb-2772-babc-18ba-28ce9a66500b%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/e0583c38-4824-6ef2-a9de-9dd8f84f5015%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/2840b03e-7290-f633-2107-c2c49fc191de%28Office.15%29.aspx)|
|[Prefix](http://msdn.microsoft.com/library/82263675-b1c4-7190-784a-1741c70329c1%28Office.15%29.aspx)|
|[ReplyStyle](http://msdn.microsoft.com/library/bb5e0d3d-29ca-33dd-b437-cf2526451352%28Office.15%29.aspx)|
|[ResponseStyle](http://msdn.microsoft.com/library/6c20276c-51c1-3164-a28f-ac415c911cbb%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/cfe619d2-3a7e-c8af-de17-be2363de0a56%28Office.15%29.aspx)|
|[ShowOn](http://msdn.microsoft.com/library/62646ba1-7e25-8402-5530-d62fe45503e5%28Office.15%29.aspx)|

## See also


#### Other resources


[Action Object Members](http://msdn.microsoft.com/library/b423cdd8-c67e-a53b-9166-eacfd5a33e7c%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
