---
title: PropertyPages Object (Outlook)
keywords: vbaol11.chm160
f1_keywords:
- vbaol11.chm160
ms.prod: outlook
api_name:
- Outlook.PropertyPages
ms.assetid: 9850ae7b-f167-d3b2-2e9b-f1df1e4922ec
ms.date: 06/08/2017
---


# PropertyPages Object (Outlook)

Contains the custom property pages that have been added to the Microsoft Outlook **Options** dialog box or to the folder **Properties** dialog box.


## Remarks

You receive a  **PropertyPages** object as a parameter of the **[OptionsPagesAdd](application-optionspagesadd-event-outlook.md)** event. Use the **[Add](propertypages-add-method-outlook.md)** method to add a **[PropertyPage](propertypage-object-outlook.md)** object to the **PropertyPages** object.


 **Note**  If more than one program handles the  **OptionsPagesAdd** event, the order in which the programs receive the event (and therefore, the order in which pages are added to the **PropertyPages** object) cannot be guaranteed.


## Methods



|**Name**|
|:-----|
|[Add](propertypages-add-method-outlook.md)|
|[Item](propertypages-item-method-outlook.md)|
|[Remove](propertypages-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](propertypages-application-property-outlook.md)|
|[Class](propertypages-class-property-outlook.md)|
|[Count](propertypages-count-property-outlook.md)|
|[Parent](propertypages-parent-property-outlook.md)|
|[Session](propertypages-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
