---
title: Pages Object (Access)
keywords: vbaac10.chm10125
f1_keywords:
- vbaac10.chm10125
ms.prod: access
api_name:
- Access.Pages
ms.assetid: e77c8d31-1cb7-d647-6faa-2eb234ce0708
ms.date: 06/08/2017
---


# Pages Object (Access)

The  **Pages** collection contains all **[Page](page-object-access.md)** objects in a tab control.


## Remarks

The  **Pages** collection is a special kind of **Controls** collection belonging to the tab control. It contains **Page** objects, which are controls. The **Pages** collection differs from a typical **Controls** collection in that you can add and remove **Page** objects by using methods of the **Pages** collection.

To add a new  **Page** object to the **Pages** collection from Visual Basic, use the **[Add](pages-add-method-access.md)** method of the **Pages** collection. To remove an existing **Page** object, use the **[Remove](pages-remove-method-access.md)** method of the **Pages** collection. To count the number of **Page** objects in the **Pages** collection, use the **[Count](pages-count-property-access.md)** property of the **Pages** collection.

You can also use the  **[CreateControl](application-createcontrol-method-access.md)** method to add a **Page** object to the **Pages** collection of a tab control. To do this, you must specify the name of the tab control for the _Parent_ argument of the **CreateControl** function. The **[ControlType](page-controltype-property-access.md)** property constant for a **Page** object is **acPage**.

You can enumerate through the  **Pages** collection by using the **For Each...Next** statement.

Individual  **Page** objects in the **Pages** collection are indexed beginning with zero.


## Methods



|**Name**|
|:-----|
|[Add](pages-add-method-access.md)|
|[Remove](pages-remove-method-access.md)|

## Properties



|**Name**|
|:-----|
|[Count](pages-count-property-access.md)|
|[Item](pages-item-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
