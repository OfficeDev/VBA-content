---
title: TableView.MultiLine Property (Outlook)
keywords: vbaol11.chm2524
f1_keywords:
- vbaol11.chm2524
ms.prod: outlook
api_name:
- Outlook.TableView.Multiline
ms.assetid: 732b39ca-ec7f-5a43-db55-3351a368b599
ms.date: 06/08/2017
---


# TableView.MultiLine Property (Outlook)

Returns or sets an  **[OlMultiLine](olmultiline-enumeration-outlook.md)** constant that determines how multiple lines are displayed in the **[TableView](tableview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **Multiline**

 _expression_ A variable that represents a **TableView** object.


## Remarks

If the value of the  **[AutomaticColumnSizing](tableview-automaticcolumnsizing-property-outlook.md)** property is set to **False** or if the value of the **[AllowInCellEditing](tableview-allowincellediting-property-outlook.md)** property is set to **True** , the value of this property is automatically set to **olAlwaysSingleLine** .


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

