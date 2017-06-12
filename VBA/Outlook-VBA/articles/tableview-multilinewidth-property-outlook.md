---
title: TableView.MultiLineWidth Property (Outlook)
keywords: vbaol11.chm2525
f1_keywords:
- vbaol11.chm2525
ms.prod: outlook
api_name:
- Outlook.TableView.MultiLineWidth
ms.assetid: 4b2a7d06-f6f7-fa9f-8957-bdc451e248e7
ms.date: 06/08/2017
---


# TableView.MultiLineWidth Property (Outlook)

Returns or sets a  **Long** value that represents the text width (in characters) needed to trigger multiline mode in the **[TableView](tableview-object-outlook.md)** object . Read/write


## Syntax

 _expression_ . **MultiLineWidth**

 _expression_ A variable that represents a **TableView** object.


## Remarks

This property can be set to a value between 1 and 999. If this property is set to a value less than 1, the property is set to 1. If this property is set to a value greater than 999, the property is set to 999. The default value for this property is 100.

This property only applies if the  **[Multiline](tableview-multiline-property-outlook.md)** property of the **TableView** object is set to **olWidthMultiLine** . The value of this property determines the point at which the **TableView** object displays text for an Outlook item in multiline mode.


## Example

The following Visual Basic for Applications (VBA) example configures the current  **TableView** object so that, if text in the view is longer than 50 characters, the text is displayed in multiline mode. The **MultiLine** property cannot be set to **olWidthMultiLine** unless the **[AutomaticColumnSizing](tableview-automaticcolumnsizing-property-outlook.md)** property is set to **True** .


```vb
Private Sub ConfigureMultiLineView() 
 
 Dim objTableView As TableView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 With objTableView 
 
 ' Set the TableView object so that, 
 
 ' if the text in the view is longer 
 
 ' than 50 characters, the text is 
 
 ' displayed in multiline mode. 
 
 .AutomaticColumnSizing = True 
 
 .MultiLine = olWidthMultiLine 
 
 .MultiLineWidth = 50 
 
 
 
 ' Save the table view. 
 
 .Save 
 
 End With 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

