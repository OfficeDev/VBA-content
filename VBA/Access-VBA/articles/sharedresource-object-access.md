---
title: SharedResource Object (Access)
keywords: vbaac10.chm14654
f1_keywords:
- vbaac10.chm14654
ms.prod: access
api_name:
- Access.SharedResource
ms.assetid: a97163fa-f833-ed1c-aea5-1a7bab783eba
ms.date: 06/08/2017
---


# SharedResource Object (Access)

Represents a Microsoft Office theme or image that is available as a shared resource in the database.


## Remarks

A shared recource is stored once in the database, but can be used many times. For example, you may want to display your company logo on every form that you create. In earlier versions of Access, you had to import the logo into every form. In Access, you can add the logo as a shared image. Then , it will be displayed in the  **Image Gallery** that is displayed when you click the **Insert Image** dropdown menu for the **Controls** group in the **Design** tab.

To import an image as a  **SharedResource** object, use the **[AddSharedImage](codeproject-addsharedimage-method-access.md)** method of the **[CodeProject](codeproject-object-access.md)** object or the **[AddSharedImage](currentproject-addsharedimage-method-access.md)** method of the **[CurrentProject](currentproject-object-access.md)** object.


## Methods



|**Name**|
|:-----|
|[Delete](sharedresource-delete-method-access.md)|

## Properties



|**Name**|
|:-----|
|[Name](sharedresource-name-property-access.md)|
|[Parent](sharedresource-parent-property-access.md)|
|[Type](sharedresource-type-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
