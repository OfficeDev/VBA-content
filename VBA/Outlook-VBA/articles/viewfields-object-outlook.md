---
title: ViewFields Object (Outlook)
keywords: vbaol11.chm3184
f1_keywords:
- vbaol11.chm3184
ms.prod: outlook
api_name:
- Outlook.ViewFields
ms.assetid: 2516faed-ed11-6cb3-ce9c-b6afa788e909
ms.date: 06/08/2017
---


# ViewFields Object (Outlook)

Represents the collection of  **[ViewField](viewfield-object-outlook.md)** objects in a view.


## Remarks

The  **ViewFields** collection represents the Outlook item properties available for display in the view. Use the **[Add](viewfields-add-method-outlook.md)** method of the **ViewFields** collection to add a view field for the following objects derived from the **[View](view-object-outlook.md)** object:


-  **[CardView](cardview-object-outlook.md)**
    
-  **[TableView](tableview-object-outlook.md)**
    
In a table view, the order of  **ViewField** objects in the **ViewFields** collection is not the same as the order that field columns are displayed in the table view. A workaround to obtain the column order is to parse the string returned by the **[View.XML](view-xml-property-outlook.md)** property.


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


## Methods



|**Name**|
|:-----|
|[Add](viewfields-add-method-outlook.md)|
|[Insert](viewfields-insert-method-outlook.md)|
|[Item](viewfields-item-method-outlook.md)|
|[Remove](viewfields-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](viewfields-application-property-outlook.md)|
|[Class](viewfields-class-property-outlook.md)|
|[Count](viewfields-count-property-outlook.md)|
|[Parent](viewfields-parent-property-outlook.md)|
|[Session](viewfields-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
