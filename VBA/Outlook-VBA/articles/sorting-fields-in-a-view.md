---
title: Sorting Fields in a View
ms.prod: outlook
ms.assetid: 23d87740-12eb-aa00-1cf5-4dfa5895722d
ms.date: 06/08/2017
---


# Sorting Fields in a View

 Outlook items can be sorted by adding one or more Outlook item properties to the **[OrderFields](orderfields-object-outlook.md)** collection of any of the following objects:


-  **[BusinessCardView](businesscardview-object-outlook.md)**
    
-  **[CardView](cardview-object-outlook.md)**
    
-  **[IconView](iconview-object-outlook.md)**
    
-  **[TableView](tableview-object-outlook.md)**
    

Outlook items in a  **[CalendarView](calendarview-object-outlook.md)** or **[TimelineView](timelineview-object-outlook.md)** object are displayed in chronological order, depending on the values of the Outlook item properties specified for the **StartField** and **EndField** properties of the view.

The  **OrderFields** collection for those views can be retrieved by calling the **SortFields** property of the appropriate view object. The **[Add](orderfields-add-method-outlook.md)** method of the **OrderFields** collection is used to create an **[OrderField](orderfield-object-outlook.md)** object that represents the Outlook item property to be sorted.

## Specifying Properties for Sorting

You can add either built-in or custom Outlook item properties to the  **OrderFields** collection. The order in which the properties are included in the **OrderFields** collection determines the order in which the properties are sorted, while the **[IsDescending](orderfield-isdescending-property-outlook.md)** property of the **OrderField** object which represents an Outlook item property determines whether the values of that property are sorted in ascending or descending order.


## Specifying Built-In Properties for Sorting

The following guidelines should be used when specifying built-in Outlook item properties:


- Built-in properties can be specified either by property name (for example, "Subject") or by namespace (for example, "http://schemas.microsoft.com/mapi/proptag/0x0037001E").
    
- Property names are not case-sensitive and cannot include spaces.
    
- Namespace identifiers are case-sensitive, must follow URL encoding rules, and cannot be enclosed in square brackets ([]).
    

## Specifying Custom Properties for Sorting

The following guidelines should be used when specifying custom properties:


- The custom property must be available in the  **[UserDefinedProperties](userdefinedproperties-object-outlook.md)** collection for the parent **[Folder](folder-object-outlook.md)** object.
    
- Custom properties should be specified by property name (for example, "[Shoe Size]").
    
- Custom property names are not case-sensitive, can include spaces, and should be enclosed in square brackets ([]) if they contain spaces.
    
For more information about property identifiers, see  [Properties Overview](properties-overview.md).


