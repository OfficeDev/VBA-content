---
title: Column Object (Outlook)
keywords: vbaol11.chm3191
f1_keywords:
- vbaol11.chm3191
ms.prod: outlook
api_name:
- Outlook.Column
ms.assetid: b7eb6916-2d80-57c3-2077-47a2a4c73185
ms.date: 06/08/2017
---


# Column Object (Outlook)

Represents a column of data in a  **[Table](table-object-outlook.md)** object.


## Remarks

A  **Table** is composed of rows and columns. It represents a read-only dynamic rowset of data in a **[Folder](folder-object-outlook.md)** or **[Search](search-object-outlook.md)** object. You can regard each row of a **Table** as an item in the folder, each column as a property of the item. By default, a **Table** contains only a subset of properties for items in the folder. This makes the **Table** an in-memory lightweight rowset that allows fast enumeration and filtering of items in the folder.

To obtain the value of a property (column) for a specific item (row) in a  **Table**, you can either use the **[Table.GetArray](table-getarray-method-outlook.md)** method and index into the returned array, or use the **[Row.Item](row-item-method-outlook.md)** method, specifying the **[Name](column-name-property-outlook.md)** of the column.


## Properties



|**Name**|
|:-----|
|[Application](column-application-property-outlook.md)|
|[Class](column-class-property-outlook.md)|
|[Name](column-name-property-outlook.md)|
|[Parent](column-parent-property-outlook.md)|
|[Session](column-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
