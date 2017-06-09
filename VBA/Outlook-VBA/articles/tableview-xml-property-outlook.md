---
title: TableView.XML Property (Outlook)
keywords: vbaol11.chm2514
f1_keywords:
- vbaol11.chm2514
ms.prod: outlook
api_name:
- Outlook.TableView.XML
ms.assetid: 0f085984-3056-6603-ca12-a4436abf429f
ms.date: 06/08/2017
---


# TableView.XML Property (Outlook)

Returns or sets a  **String** value that specifies the XML definition of the view. Read/write.


## Syntax

 _expression_ . **XML**

 _expression_ A variable that represents a **TableView** object.


## Remarks

The XML definition describes the view type by using a series of tags and keywords corresponding to various properties of the view itself. When the view is created, the XML definition is parsed to render the settings for the new view.

To determine how the XML should be structured when creating views, you can create a view by using the Outlook user interface and then you can retrieve the  **XML** property for that view.


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

