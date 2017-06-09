---
title: Filtering and Specifying Fields in a View
ms.prod: outlook
ms.assetid: 99cef6bf-1bf6-706f-9972-22005a255416
ms.date: 06/08/2017
---


# Filtering and Specifying Fields in a View

 You can specify which Outlook item properties are displayed in a view by adding one or more properties to the **[ViewFields](viewfields-object-outlook.md)** collection of any of the following objects:


-  **[CardView](cardview-object-outlook.md)**
    
-  **[TableView](tableview-object-outlook.md)**
    

 **[BusinessCardView](businesscardview-object-outlook.md)** , **[CalendarView](calendarview-object-outlook.md)**,  **[IconView](iconview-object-outlook.md)**, and  **[TimelineView](timelineview-object-outlook.md)** objects use other methods of determining which Outlook item properties are displayed within the view.The fields displayed for the **BusinessCardView** object, for example, are determined by the Electronic Business Card (EBC) layout associated with each displayed Outlook item.

The  **ViewFields** collection for those views can be retrieved by calling the **ViewFields** property of the appropriate view object. The **[Add](viewfields-add-method-outlook.md)** method of the **ViewFields** collection is used to create a **[ViewField](viewfield-object-outlook.md)** object that represents the Outlook item property to be displayed in the view.
A  **ViewField** object not only identifies an Outlook item property to display within the view, but also describes how the values for that property should be displayed. You can change how properties are displayed in a view.

## Filtering Outlook Items

Outlook items can be filtered in any view derived from the  **[View](view-object-outlook.md)** object by specifying a DAV Searching and Locating (DASL) filter expression in the **[Filter](view-filter-property-outlook.md)** property of the **View** object. For more information about creating a DASL filter expression with which to filter Outlook items, see [Filtering Items](filtering-items.md).


