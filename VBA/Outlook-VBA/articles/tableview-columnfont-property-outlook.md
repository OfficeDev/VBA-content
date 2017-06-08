---
title: TableView.ColumnFont Property (Outlook)
keywords: vbaol11.chm2534
f1_keywords:
- vbaol11.chm2534
ms.prod: outlook
api_name:
- Outlook.TableView.ColumnFont
ms.assetid: f69ff872-1823-b5c0-9a3d-d4cf72973be1
ms.date: 06/08/2017
---


# TableView.ColumnFont Property (Outlook)

Returns a  **[ViewFont](viewfont-object-outlook.md)** object that represents the font used when displaying column headers in the **[TableView](tableview-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **ColumnFont**

 _expression_ A variable that represents a **TableView** object.


## Example

The following Visual Basic for Applications (VBA) sample increments the value of the  **[Size](viewfont-size-property-outlook.md)** property for the **ViewFont** object returned from the **ColumnFont** property for the current **TableView** object.


```vb
Private Sub IncreaseColumnFontSize() 
 
 Dim objTableView As TableView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Increment the Size property of the 
 
 ' ViewFont object obtained from the 
 
 ' ColumnFont property, but only 
 
 ' if the font is less than 24 points 
 
 ' in size. 
 
 If objTableView.ColumnFont.Size < 24 Then 
 
 objTableView.ColumnFont.Size = _ 
 
 objTableView.ColumnFont.Size + 1 
 
 
 
 ' Save the table view. 
 
 objTableView.Save 
 
 End If 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

