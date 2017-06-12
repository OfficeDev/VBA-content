---
title: ColumnFormat.Label Property (Outlook)
keywords: vbaol11.chm2728
f1_keywords:
- vbaol11.chm2728
ms.prod: outlook
api_name:
- Outlook.ColumnFormat.Label
ms.assetid: cf104506-3eca-6695-3d3b-05022ce6fba4
ms.date: 06/08/2017
---


# ColumnFormat.Label Property (Outlook)

Returns or sets a  **String** value that represents the column label and tooltip displayed for the property to which the **[ColumnFormat](columnformat-object-outlook.md)** object is associated. Read/write.


## Syntax

 _expression_ . **Label**

 _expression_ A variable that represents a **ColumnFormat** object.


## Remarks

For built-in Outlook properties, the default value for this property is the localized name of the property. For custom Outlook properties, the default value for this property is the name of the property.

The value of this property applies only to the tooltip for Outlook properties in which the column header is represented as an icon.


## Example

The following Visual Basic for Applications (VBA) example iterates through the  **[ViewFields](tableview-viewfields-property-outlook.md)** collection of the current **[TableView](tableview-object-outlook.md)** object, displaying the label and XML schema names of each **ViewField** object in the collection.


```vb
Private Sub DisplayTableViewFields() 
 
 Dim objTableView As TableView 
 
 Dim objViewField As ViewField 
 
 Dim strOutput As String 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Iterate through the ViewFields collection for 
 
 ' the table view, obtaining the label and the 
 
 ' XML schema name for each field included in 
 
 ' the view. 
 
 For Each objViewField In objTableView.ViewFields 
 
 With objViewField 
 
 strOutput = strOutput &; .ColumnFormat.Label &; _ 
 
 " (" &; .ViewXMLSchemaName &; ")" &; vbCrLf 
 
 End With 
 
 Next 
 
 
 
 ' Display a dialog box containing the concatenated 
 
 ' view field information. 
 
 MsgBox strOutput 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[ColumnFormat Object](columnformat-object-outlook.md)

