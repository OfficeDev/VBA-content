---
title: Enumerating, Searching, and Filtering Items in a Folder
ms.prod: outlook
ms.assetid: d786d292-7a0e-0e1a-e132-affbfde37744
ms.date: 06/08/2017
---


# Enumerating, Searching, and Filtering Items in a Folder

The Outlook object model provides several entry points that support enumerating, searching, and filtering items in a folder. 

## Enumerating Items

The  **[Items](items-object-outlook.md)**,  **[Table](table-object-outlook.md)**, and  **[Selection](selection-object-outlook.md)** objects all support enumerating items in a folder. However, there are specific scenarios where you would choose one over the others.

You can obtain an  **Items** collection by calling **[Folder.Items](folder-items-property-outlook.md)** which returns the items in the folder. Each item object in the collection is complete with all its explicit built-in properties and custom properties, and supports read-write operations. The **Items** collection also supports filters and events that fire when items are added, changed, or removed from the collection.

You can use  **[Folder.GetTable](folder-gettable-method-outlook.md)** or **[Search.GetTable](search-gettable-method-outlook.md)** to obtain a **Table** object that represents a set of items in a folder or search folder. In both cases, you can specify a filter to obtain a subset of the items in the folder, or, if you do not specify any filter, obtain all the items in the folder. By default, each item in the returned **Table** contains only a default subset of its properties. 

You can view each row of a **Table** as an item in the folder, each column as a property of the item, and the **Table** is an in-memory light-weight rowset that allows fast enumeration and filtering of items in the folder. Although additions and deletions of the underlying folder are reflected by the rows in the **Table**, the  **Table** does not support any events for adding, changing, and removing rows. 

If you require a writeable object from the **Table** row, obtain the Entry ID for that row from the default EntryID column in the **Table** and then use the **[GetItemFromID](namespace-getitemfromid-method-outlook.md)** method of the **[NameSpace](namespace-object-outlook.md)** object to obtain a full item, such as a **[MailItem](mailitem-object-outlook.md)** or **[ContactItem](contactitem-object-outlook.md)**, that supports read-write operations. For more information on default columns in a  **Table**, see  [Default Properties Displayed in a Table Object](default-properties-displayed-in-a-table-object.md).

The  **Selection** object supports enumerating items that a user has currently selected in an explorer. Since the explorer displays the contents of a folder, the **Selection** object supports enumeration of items in that folder as per the user's selection.

 **Note**  A folder in Outlook can contain heterogeneous items. For example, the Contacts folder supports creating contact items and distribution list items by default. Since the  **Items**,  **Table**, and  **Selection** objects encapsulate items in a folder or search folder, the items in them do not necessarily have the same message class. When enumerating items in these collections and objects, it is a good practice to first check for the message class of each item before accessing the item's properties.


## Searching and Filtering Items

The  **Items**,  **Table**,  **[Application](application-object-outlook.md)**, and  **[View](view-object-outlook.md)** objects support searching and filtering of items in a folder. The following table describes and compares these entry points:


| **Entry Point**| **Action**| **Object of Search Filter**| **Jet Filter Support**| **DASL Filter Support**| **Comments**|
|:-----|:-----|:-----|:-----|:-----|:-----|
| **[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)**|Sets the criteria for a  **Search** object and returns the **Search** object. **[Search.Results](search-results-property-outlook.md)** specifies the search results. **[Search.Save](search-save-method-outlook.md)** updates a search folder with the search results.|Folder|No|Yes||
| **Folder.GetTable**|Returns a  **Table** of items in a folder based on any given filter.|Folder|Yes|Yes|Certain properties are not supported in the  **Table** filter, including binary properties, and HTML or RTF body content. For more information, see [Unsupported Properties in a Table Object or Table Filter](unsupported-properties-in-a-table-object-or-table-filter.md).|
| **[Items.Find](items-find-method-outlook.md)**|Searches for first item that satisfies the specified filter. |Folder|Yes|No|Certain properties are not supported in the filter, including  **Body**. For more information, see  **Items.Find**.|
| **[Items.Restrict](items-restrict-method-outlook.md)**|Filters given set of items based on specified restrictions and returns another  **Items** collection.|Folder|Yes|Yes|Certain properties are not supported in the filter, for example,  **Body**. For more information, see  **Items.Restrict**.|
| **Search.GetTable**|Returns a  **Table** of items in a search folder based on any filter derived from **Application.AdvancedSearch**.|Search folder|No|Yes| **Search.GetTable** derives its filter from the **Search** object (specifically the **[Search.Filter](search-filter-property-outlook.md)** property) returned from **Application.AdvancedSearch**.|
| **[Table.Restrict](table-restrict-method-outlook.md)**|Filters rows in the given table based on a specified filter and returns another  **Table** object.|Folder|Yes|Yes|Certain properties are not supported in the  **Table** filter, including binary properties, and HTML or RTF body content. For more information, see [Unsupported Properties in a Table Object or Table Filter](unsupported-properties-in-a-table-object-or-table-filter.md).|
| **[View.Filter](view-filter-property-outlook.md)**|Sets a view's filter without changing the view's XML. Setting the filter for a view only changes the view in the user interface and does not result in a filtered  **Items** collecton.|Folder|No|Yes||
| **[View.XML](view-xml-property-outlook.md)**|Gets or sets the XML for a view. Modifying the <Filter> node changes the view's filter. Setting the XML for a view only changes the view in the user interface and does not result in a filtered  **Items** collecton.|Folder|No|Yes|View XML is being deprecated. Use the View object model to program views.|



