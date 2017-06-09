---
title: SyncObjects Object (Outlook)
keywords: vbaol11.chm94
f1_keywords:
- vbaol11.chm94
ms.prod: outlook
api_name:
- Outlook.SyncObjects
ms.assetid: 88e59f63-d834-b174-bbda-0af0cf2d0520
ms.date: 06/08/2017
---


# SyncObjects Object (Outlook)

Contains a set of  **[SyncObject](syncobject-object-outlook.md)** objects representing the **Send/Receive** groups for a user.


## Remarks

Use the  **[SyncObjects](namespace-syncobjects-property-outlook.md)** property to return the **SyncObjects** object from a **[NameSpace](namespace-object-outlook.md)** object.

The  **SyncObjects** object is read-only. You cannot add an item to the collection. However, note that you can add one **Send/Receive** group using the **AppFolders** property which will create a **Send/Receive** group called Application Folders.


## Example

The following Microsoft Visual Basic for Applications (VBA) example retrieves the  **SyncObjects** object for the MAPI **NameSpace** object.


```
Set mySyncObjects = Application.GetNameSpace("MAPI").SyncObjects
```


## Methods



|**Name**|
|:-----|
|[Item](syncobjects-item-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[AppFolders](syncobjects-appfolders-property-outlook.md)|
|[Application](syncobjects-application-property-outlook.md)|
|[Class](syncobjects-class-property-outlook.md)|
|[Count](syncobjects-count-property-outlook.md)|
|[Parent](syncobjects-parent-property-outlook.md)|
|[Session](syncobjects-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
