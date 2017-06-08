---
title: TableView.ShowNewItemRow Property (Outlook)
keywords: vbaol11.chm2527
f1_keywords:
- vbaol11.chm2527
ms.prod: outlook
api_name:
- Outlook.TableView.ShowNewItemRow
ms.assetid: 2e389bb6-9d1f-6c9d-0cdc-b177705d620b
ms.date: 06/08/2017
---


# TableView.ShowNewItemRow Property (Outlook)

Returns or sets a  **Boolean** value that determines if the new item row is displayed in the **[TableView](tableview-object-outlook.md)** object. Read/write


## Syntax

 _expression_ . **ShowNewItemRow**

 _expression_ A variable that represents a **TableView** object.


## Remarks

The value of this property applies only if the  **[AllowInCellEditing](tableview-allowincellediting-property-outlook.md)** property is set to **True** .


## Example

The following Visual Basic for Applications (VBA) example configures the current  **TableView** object so that in-cell editing is allowed and the new item row is displayed in the view.


```vb
Private Sub ConfigureEditableView() 
 
 Dim objTableView As TableView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 With objTableView 
 
 ' ShowNewItemRow is ignored if 
 
 ' AllowInCellEditing is set to 
 
 ' False. 
 
 .AllowInCellEditing = True 
 
 
 
 ' Display the new item row in 
 
 ' the table view. 
 
 .ShowNewItemRow = True 
 
 
 
 ' Save the table view. 
 
 .Save 
 
 End With 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

