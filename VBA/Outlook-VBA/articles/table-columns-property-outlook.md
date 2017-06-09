---
title: Table.Columns Property (Outlook)
keywords: vbaol11.chm2236
f1_keywords:
- vbaol11.chm2236
ms.prod: outlook
api_name:
- Outlook.Table.Columns
ms.assetid: 57005ab1-ad49-296d-5b34-24dfd8f0987f
ms.date: 06/08/2017
---


# Table.Columns Property (Outlook)

Returns a  **[Columns](columns-object-outlook.md)** collection object that contains the columns defined for the **[Table](table-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **Columns**

 _expression_ A variable that represents a **Table** object.


## Remarks

The  **Columns** collection object is the default member of the **Table** object.

While rows in a  **Table** correspond to items in the parent **[Folder](folder-object-outlook.md)** or **[Search](search-object-outlook.md)** object of the **Table** , **Columns** in a **Table** correspond to the properties of these items. Default columns are defined for all folders depending on the parent folder of the **Table** object. For example, the default properties for the Inbox are: **EntryID** , **Subject** , **CreationTime** , **LastModificationTime** , and **MessageClass** . For more information on default properties for a **Table** , see[Default Properties Displayed in a Table Object](http://msdn.microsoft.com/library/649c64f3-2d1e-23f1-bf13-3368da79e62b%28Office.15%29.aspx).

To add  **[Column](column-object-outlook.md)** objects to the **Columns** collection of a **Table** , use **[Columns.Add](columns-add-method-outlook.md)** . To remove the default column set, use **[Columns.RemoveAll](columns-removeall-method-outlook.md)** . For more information on adjusting columns of a **Table** , see[Adding Columns to a Table Object](http://msdn.microsoft.com/library/c1d652ef-8082-70f3-1216-d39e976e6b21%28Office.15%29.aspx).


## See also


#### Concepts


[Table Object](table-object-outlook.md)

