---
title: ColumnFormat Object (Outlook)
keywords: vbaol11.chm3189
f1_keywords:
- vbaol11.chm3189
ms.prod: outlook
api_name:
- Outlook.ColumnFormat
ms.assetid: acbbdd97-e695-d1e7-c7ba-24f75efbf22c
ms.date: 06/08/2017
---


# ColumnFormat Object (Outlook)

Represents the display properties of an order field or view field in a view.


## Remarks

The  **ColumnFormat** object represents the display properties, such as the alignment or field type, of an **[OrderField](orderfield-object-outlook.md)** or **[ViewField](viewfield-object-outlook.md)** object. Use the **[ColumnFormat](viewfield-columnformat-property-outlook.md)** property of the **ViewField** object to access the display properties of a view field.

Use the  **[Label](columnformat-label-property-outlook.md)** property to obtain or change the text used to label the field, or the **[Align](columnformat-align-property-outlook.md)** property to determine the alignment of the contents within the field.

Use the  **[FieldType](columnformat-fieldtype-property-outlook.md)** property to determine the type and form of the data displayed for that field, and the **[FieldFormat](columnformat-fieldformat-property-outlook.md)** property to determine how to format the data for that field.


## Example

The following Visual Basic for Applications (VBA) example iterates through the  **[ViewFields](tableview-viewfields-property-outlook.md)** collection of the current **[TableView](tableview-object-outlook.md)** object, displaying the label and XML schema names of each **ViewField** object in the collection.


```
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
 
 strOutput = strOutput &amp; .ColumnFormat.Label &amp; _ 
 
 " (" &amp; .ViewXMLSchemaName &amp; ")" &amp; vbCrLf 
 
 End With 
 
 Next 
 
 
 
 ' Display a dialog box containing the concatenated 
 
 ' view field information. 
 
 MsgBox strOutput 
 
 End If 
 
End Sub 
 

```


## Properties



|**Name**|
|:-----|
|[Align](columnformat-align-property-outlook.md)|
|[Application](columnformat-application-property-outlook.md)|
|[Class](columnformat-class-property-outlook.md)|
|[FieldFormat](columnformat-fieldformat-property-outlook.md)|
|[FieldType](columnformat-fieldtype-property-outlook.md)|
|[Label](columnformat-label-property-outlook.md)|
|[Parent](columnformat-parent-property-outlook.md)|
|[Session](columnformat-session-property-outlook.md)|
|[Width](columnformat-width-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
