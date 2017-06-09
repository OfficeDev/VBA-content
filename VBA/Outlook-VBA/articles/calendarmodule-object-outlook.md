---
title: CalendarModule Object (Outlook)
keywords: vbaol11.chm3194
f1_keywords:
- vbaol11.chm3194
ms.prod: outlook
api_name:
- Outlook.CalendarModule
ms.assetid: 9203024d-9cef-75e0-600f-f3899e24761a
ms.date: 06/08/2017
---


# CalendarModule Object (Outlook)

Represents the  **Calendar** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **CalendarModule** object, derived from the **[NavigationModule](http://msdn.microsoft.com/library/76565eaf-1e64-f5d4-b90f-ba156863802c%28Office.15%29.aspx)** object, provides access to the navigation groups contained in the **Calendar** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](http://msdn.microsoft.com/library/7c1a1313-94a4-fa68-7e70-66d85496fec0%28Office.15%29.aspx)** method or the **[Item](http://msdn.microsoft.com/library/ee8fdd9c-2b94-29c3-7622-f6e5c8c5399c%28Office.15%29.aspx)** method of the **[Modules](http://msdn.microsoft.com/library/f7311738-369c-4dd6-947c-9382195bc944%28Office.15%29.aspx)** collection for the parent **[NavigationPane](http://msdn.microsoft.com/library/b6538c72-6115-99fc-c926-e0532a747823%28Office.15%29.aspx)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](http://msdn.microsoft.com/library/ee1fc78a-9720-c8d0-964c-0178ddbe8af6%28Office.15%29.aspx)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleCalendar**, you can then cast the **NavigationModule** object reference as a **CalendarModule** object to access the **[NavigationGroups](http://msdn.microsoft.com/library/2f19eceb-24e6-a55c-7013-c840bd0c9fbb%28Office.15%29.aspx)** property for that navigation module.

You can use the  **[Visible](http://msdn.microsoft.com/library/e34a7247-59aa-0a7f-fe8c-b439f683b22c%28Office.15%29.aspx)** property to determine if the navigation module is visible and the **[Position](http://msdn.microsoft.com/library/3857d981-acd7-975c-0ff1-453ee2b7402e%28Office.15%29.aspx)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](http://msdn.microsoft.com/library/1c1e262e-8775-5039-a9f2-1a279f4289a9%28Office.15%29.aspx)** property to return the display name of the **Calendar** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/9bc00532-e527-aa34-7377-704e1d54a2ff%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/0f106f3d-b4c4-54ce-746e-89cd5cac62e7%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/1c1e262e-8775-5039-a9f2-1a279f4289a9%28Office.15%29.aspx)|
|[NavigationGroups](http://msdn.microsoft.com/library/2f19eceb-24e6-a55c-7013-c840bd0c9fbb%28Office.15%29.aspx)|
|[NavigationModuleType](http://msdn.microsoft.com/library/cb63445b-0438-c97e-0b38-eaf17b6b739e%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1a0637c3-e449-32ba-8597-87b8a04235f4%28Office.15%29.aspx)|
|[Position](http://msdn.microsoft.com/library/3857d981-acd7-975c-0ff1-453ee2b7402e%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/df23c975-9ac9-4ed9-0369-dce6b59e518a%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/e34a7247-59aa-0a7f-fe8c-b439f683b22c%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[CalendarModule Object Members](http://msdn.microsoft.com/library/82731a1f-3ebe-1cb0-9e8b-d370a0b8f954%28Office.15%29.aspx)
