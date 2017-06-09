---
title: ViewField.ColumnFormat Property (Outlook)
keywords: vbaol11.chm2544
f1_keywords:
- vbaol11.chm2544
ms.prod: outlook
api_name:
- Outlook.ViewField.ColumnFormat
ms.assetid: 0014f1d8-5380-3301-558a-7fd8d49afff9
ms.date: 06/08/2017
---


# ViewField.ColumnFormat Property (Outlook)

Returns a  **[ColumnFormat](columnformat-object-outlook.md)** object that represents the formatting information for the **[ViewField](viewfield-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **ColumnFormat**

 _expression_ A variable that represents a **ViewField** object.


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


[ViewField Object](viewfield-object-outlook.md)

