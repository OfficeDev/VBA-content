---
title: OrderField.ViewXMLSchemaName Property (Outlook)
keywords: vbaol11.chm2687
f1_keywords:
- vbaol11.chm2687
ms.prod: outlook
api_name:
- Outlook.OrderField.ViewXMLSchemaName
ms.assetid: a88c22ff-3d30-a4f2-87f6-6c32c1c2acb7
ms.date: 06/08/2017
---


# OrderField.ViewXMLSchemaName Property (Outlook)

Returns a  **String** value that represents the XML schema name for the property referenced by the **[OrderField](orderfield-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **ViewXMLSchemaName**

 _expression_ A variable that represents an **OrderField** object.


## Remarks

The value of this property contains the name of the property as it is included within the XML definition of the view containing the  **ViewField** object. This value may not match the name used to refer to the property when the **OrderField** object was defined.


## Example

The following Visual Basic for Applications (VBA) example iterates through the  **[SortFields](tableview-sortfields-property-outlook.md)** collection of the current **[TableView](tableview-object-outlook.md)** object, displaying the label and XML schema names of each **OrderField** object in the collection.


```vb
Private Sub DisplayTableViewSortFields() 
 
 Dim objTableView As TableView 
 
 Dim objOrderField As OrderField 
 
 Dim strOutput As String 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Iterate through the OrderFields collection for 
 
 ' the table view, obtaining the label and the 
 
 ' XML schema name for each field used to sort 
 
 ' the items in the view. 
 
 For Each objOrderField In objTableView.SortFields 
 
 With objOrderField 
 
 strOutput = strOutput &; .ColumnFormat.Label &; _ 
 
 " (" &; .ViewXMLSchemaName &; ")" &; vbCrLf 
 
 End With 
 
 Next 
 
 
 
 ' Display a dialog box containing the concatenated 
 
 ' sort field information. 
 
 MsgBox strOutput 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[OrderField Object](orderfield-object-outlook.md)

