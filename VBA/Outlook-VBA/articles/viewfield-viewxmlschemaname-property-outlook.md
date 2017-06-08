---
title: ViewField.ViewXMLSchemaName Property (Outlook)
keywords: vbaol11.chm2543
f1_keywords:
- vbaol11.chm2543
ms.prod: outlook
api_name:
- Outlook.ViewField.ViewXMLSchemaName
ms.assetid: 69490353-b470-6092-0b8e-b0f1c1549f7a
ms.date: 06/08/2017
---


# ViewField.ViewXMLSchemaName Property (Outlook)

Returns a  **String** value that represents the XML schema name for the property referenced by the **[ViewField](viewfield-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **ViewXMLSchemaName**

 _expression_ A variable that represents a **ViewField** object.


## Remarks

The value of this property contains the name of the property as it is included within the XML definition of the view containing the  **ViewField** object. This value may not match the name used to refer to the property when the **ViewField** object was defined.


## Example

The following Visual Basic for Applications (VBA) example iterates through the  **[ViewFields](viewfields-object-outlook.md)** collection of the current **[TableView](tableview-object-outlook.md)** object, displaying the label and XML schema names of each **ViewField** object in the collection.


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

