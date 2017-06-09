---
title: Cancel an Event
ms.prod: access
ms.assetid: f91f4f8a-99fa-dca7-576a-11c76d6ddc93
ms.date: 06/08/2017
---


# Cancel an Event

Under some circumstances, you may want to include code in an event procedure that cancels the associated event. For example, you may want to include code that cancels the  **[Open](form-open-event-access.md)** event in an **Open** event procedure for a form, preventing the form from opening if certain conditions are not met.

You can cancel the following events:

||
|:-----|
|**ApplyFilter**|
|**BeforeDelConfirm**|
|**BeforeInsert**|
|**BeforeRender**|
|**BeforeUpdate**|
|**CommandBeforeExecute**|
|**DblClick**|
|**Delete**|
|**Dirty**|
|**Exit**|
|**Filter**|
|**NoData**|
|**Open**|
|**Undo**|
|**Unload**|
You cancel an event by setting an event procedure's  _Cancel_ argument to **True**.

