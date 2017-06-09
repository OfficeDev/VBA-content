---
title: CardView.XML Property (Outlook)
keywords: vbaol11.chm2594
f1_keywords:
- vbaol11.chm2594
ms.prod: outlook
api_name:
- Outlook.CardView.XML
ms.assetid: a2be1e50-ae77-785c-0dc3-2b56c3a47cc7
ms.date: 06/08/2017
---


# CardView.XML Property (Outlook)

Returns or sets a  **String** value that specifies the XML definition of the view. Read/write.


## Syntax

 _expression_ . **XML**

 _expression_ A variable that represents a **CardView** object.


## Remarks

The XML definition describes the view type by using a series of tags and keywords corresponding to various properties of the view itself. When the view is created, the XML definition is parsed to render the settings for the new view.

To determine how the XML should be structured when creating views, you can create a view by using the Outlook user interface and then you can retrieve the  **XML** property for that view.


## See also


#### Concepts


[CardView Object](cardview-object-outlook.md)

