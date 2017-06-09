---
title: PropertyPageSite Object (Outlook)
keywords: vbaol11.chm384
f1_keywords:
- vbaol11.chm384
ms.prod: outlook
api_name:
- Outlook.PropertyPageSite
ms.assetid: cdec4b4c-14b3-de0a-52c8-d5af46f4644a
ms.date: 06/08/2017
---


# PropertyPageSite Object (Outlook)

Represents the container of a custom property page.


## Remarks

Use the  **Parent** property of the ActiveX control that implements the **[PropertyPage](propertypage-object-outlook.md)** object associated with the **PropertyPageSite** object to return the **PropertyPageSite** object. The Declarations section of the module implementing the **PropertyPage** object must contain a declaration similar to the following.


```
Private myPropertyPageSite As Outlook.PropertyPageSite
```

The object is then returned from the  **Parent** property.




```
Set myPropertyPageSite = Parent
```

Use the  **[OnStatusChange](propertypagesite-onstatuschange-method-outlook.md)** method to notify Microsoft Outlook that the property page has changed.


## Methods



|**Name**|
|:-----|
|[OnStatusChange](propertypagesite-onstatuschange-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](propertypagesite-application-property-outlook.md)|
|[Class](propertypagesite-class-property-outlook.md)|
|[Parent](propertypagesite-parent-property-outlook.md)|
|[Session](propertypagesite-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
