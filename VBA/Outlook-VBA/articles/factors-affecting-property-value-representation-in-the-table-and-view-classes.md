---
title: Factors Affecting Property Value Representation in the Table and View Classes
ms.prod: outlook
ms.assetid: 13cf9945-a9e0-bb32-a2cb-74366a365ae1
ms.date: 06/08/2017
---


# Factors Affecting Property Value Representation in the Table and View Classes

There are a couple of factors that affect the type and format of a property in a  **[Table](table-object-outlook.md)** and in a **[View](view-object-outlook.md)**. String properties are affected by the store provider, and binary, date, and multi-valued properties are affected by the way the property is referenced when it is first added to a  **Table**, an  **[OrderFields](orderfields-object-outlook.md)** collection, or a **[ViewFields](viewfields-object-outlook.md)** collection, or specified as a **StartField** or **EndField** in a **[CalendarView](calendarview-object-outlook.md)** or **[TimelineView](timelineview-object-outlook.md)**.


## String Properties Affected by Store Providers

The length of the value of a string property depends on the store provider. For Exchange and OST/PST stores, the length of the string value will not exceed 255 bytes. This means that string values longer than 255 bytes will be truncated at the first 255 characters. 

For example, if you use  **[Columns.Add](columns-add-method-outlook.md)** to add the **PR_INTERNET_TRANSPORT_HEADERS** property (referenced by namespace as http://schemas.microsoft.com/mapi/proptag/0x007d001e) to a **Table**, the  **Table** will only store the first 255 characters of the full content of the property. If you need to determine the full content of the property, you must use the corresponding item's Entry ID in **[NameSpace.GetItemFromID](namespace-getitemfromid-method-outlook.md)** to obtain a full item. Once you have the item, you can use the **[PropertyAccessor](propertyaccessor-object-outlook.md)** to obtain the complete property value.


## Date, Binary, and Multi-valued Properties Affected by Property Reference

The type and format of a binary, date, or multi-valued property are affected by how the property is referenced when it is first added to a  **Table** or as a field to a **View**. Is the property referenced by its explicit built-in name (if it has one), or is it referenced by namespace (regardless of the existence of an explicit built-in name)? The following table summarizes the difference in the property value representation (in terms of type and format) per original property type:


|**Type of Property**|**Type/Format Stored**|**Type/Format Stored**|
|:-----|:-----|:-----|
||(if Property Added by Referencing an Explicit Built-in Name)|(if Property Added by Referencing a Namespace)|
|Binary|String|Byte array|
|Date|Local time|Coordinated Universal Time (UTC)|
|Multi-valued|String containing comma-separated values|1-dimensional array containing one element for each keyword|

 **Note**  For more information on referencing properties by namespace, see  [Referencing Properties by Namespace](referencing-properties-by-namespace.md).


