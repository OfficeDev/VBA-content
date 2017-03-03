---
title: FileTypes Object (Office)
keywords: vbaof11.chm257000
f1_keywords:
- vbaof11.chm257000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.FileTypes
ms.assetid: 5e8b5240-5ebd-704d-72e6-1f4ad951dfdc
---


# FileTypes Object (Office)

A collection of values of the type  **msoFileType** that determine which types of files are returned during a search.


## Remarks

There is only one  **FileTypes** collection for all searches so it's important to clear the **FileTypes** collection before executing a search unless you wish to search for file types from previous searches. The easiest way to clear the collection is to set the **FileType** property to the first file type for which you want to search. You can also remove individual types using the **Remove** method. To determine the file type of each item in the collection, use the **Item** method to return the **msoFileType** value.


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/4febf3e9-8ed5-b92b-ae0c-e5f804b27039%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/1c2d55c5-9f57-e9aa-f145-3ff61c69fb69%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/fcb569ba-c8ad-f9df-f943-b2d678f90cda%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/e286f224-9186-6198-717e-30604829287c%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/c3e9d104-e60b-4b8b-eb1c-95553dcefd89%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/89a9a9b1-1161-9dff-84db-064fc45aa022%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[FileTypes Object Members](http://msdn.microsoft.com/library/c2ecfe17-b2bb-23ef-1c2b-e5b8b5ff4fe1%28Office.15%29.aspx)
