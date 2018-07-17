---
title: NavigationFolder.IsSideBySide Property (Outlook)
keywords: vbaol11.chm2906
f1_keywords:
- vbaol11.chm2906
ms.prod: outlook
api_name:
- Outlook.NavigationFolder.IsSideBySide
ms.assetid: 00a49ce6-ad74-1f24-2aaa-e79a3409c9c9
ms.date: 06/08/2017
---


# NavigationFolder.IsSideBySide Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether the **[NavigationFolder](navigationfolder-object-outlook.md)** object is displayed in side-by-side or overlay mode. Read/write.


## Syntax

 _expression_ . **IsSideBySide**

 _expression_ A variable that represents a **NavigationFolder** object.


## Remarks

Setting this property to  **True** displays the **NavigationFolder** in side-by-side mode; otherwise, overlay mode is used to display the navigation folder in the current view of the active explorer. The default value for this property is **True** .

Setting this property has no effect for a  **NavigationFolder** object that is not associated with a **Calendar** module. If the **NavigationFolder** object is associated with a **Calendar** module, the value of this property is dependent on the following conditions:

If the  **[IsSelected](navigationfolder-isselected-property-outlook.md)** property of the **NavigationFolder** object is set to **False** , then this property value has no effect until the **IsSelected** property is set to **True** . If the **IsSelected** property is set to **True** , then the property value is applied when the **NavigationFolder** is displayed.

However, the  **IsSideBySide** property is automatically set to **True** if the **IsSelected** property for only one **NavigationFolder** associated with the parent **[CalendarModule](calendarmodule-object-outlook.md)** object is set to **True** . In other words, if the **NavigationFolder** object is the only navigation folder displayed in the current view of the active explorer, then the **IsSideBySide** property for that one **NavigationFolder** object is automatically set to **True** .


## See also


#### Concepts


[NavigationFolder Object](navigationfolder-object-outlook.md)

