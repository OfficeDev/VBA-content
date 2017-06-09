---
title: SyncObject Object (Outlook)
keywords: vbaol11.chm2984
f1_keywords:
- vbaol11.chm2984
ms.prod: outlook
api_name:
- Outlook.SyncObject
ms.assetid: 099865b6-767f-8022-6839-875624f284f7
ms.date: 06/08/2017
---


# SyncObject Object (Outlook)

Represents a  **Send\Receive** group for a user.


## Remarks

A  **Send\Receive** group lets users configure different synchronization scenarios, selecting which folders and which filters apply.

Use the  **[Item](syncobjects-item-method-outlook.md)** method to retrieve the **SyncObject** object from a **[SyncObjects](syncobjects-object-outlook.md)** object. Because the **[Name](syncobject-name-property-outlook.md)** property is the default property of the **SyncObject** object, you can identify the group by name.

The  **SyncObject** object is read-only; you cannot change its properties or create new ones. However, note that you can add one **Send/Receive** group using the **[SyncObjects.AppFolders](syncobjects-appfolders-property-outlook.md)** property which will create a **Send/Receive** group called **Application Folders**.


## Example

The following example retrieves a  **SyncObject** object by name.


```
Set mySyncObject = mySyncObjects.Item("Daily")
```


## Events



|**Name**|
|:-----|
|[OnError](syncobject-onerror-event-outlook.md)|
|[Progress](syncobject-progress-event-outlook.md)|
|[SyncEnd](syncobject-syncend-event-outlook.md)|
|[SyncStart](syncobject-syncstart-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Start](syncobject-start-method-outlook.md)|
|[Stop](syncobject-stop-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](syncobject-application-property-outlook.md)|
|[Class](syncobject-class-property-outlook.md)|
|[Name](syncobject-name-property-outlook.md)|
|[Parent](syncobject-parent-property-outlook.md)|
|[Session](syncobject-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
