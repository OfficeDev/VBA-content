---
title: SharedResources Object (Access)
keywords: vbaac10.chm14648
f1_keywords:
- vbaac10.chm14648
ms.prod: access
api_name:
- Access.SharedResources
ms.assetid: 45323141-e7df-1c70-efe2-926c1990d5e0
ms.date: 06/08/2017
---


# SharedResources Object (Access)

Represents the collection of shared resources in the database.


## Remarks

The SharedResources collection contains Microsoft Office themes and images that are stored once, but used throughout the database.

 For example, you may want to display your company logo on every form that you create. In earlier versions of Access, you had to import the logo into every form. In Access, you can add the logo as a shared image. Then , it will be displayed in the **Image Gallery** that is displayed when you click the **Insert Image** dropdown menu for the **Controls** group in the **Design** tab.

Use the  **[Resources](codeproject-resources-property-access.md)** property of the **[CodeProject](codeproject-object-access.md)** object or the **[Resources](currentproject-resources-property-access.md)** property of the **[CurrentProject](currentproject-object-access.md)** object to enumerate the **SharedResources** collection.

To import an image as a  **SharedResource** object, use the **[AddSharedImage](codeproject-addsharedimage-method-access.md)** method of the **[CodeProject](codeproject-object-access.md)** object or the **[AddSharedImage](currentproject-addsharedimage-method-access.md)** method of the **[CurrentProject](currentproject-object-access.md)** object.


## Properties



|**Name**|
|:-----|
|[Application](sharedresources-application-property-access.md)|
|[Count](sharedresources-count-property-access.md)|
|[Item](sharedresources-item-property-access.md)|
|[Parent](sharedresources-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
