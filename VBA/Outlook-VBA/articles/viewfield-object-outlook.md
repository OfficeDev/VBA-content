---
title: ViewField Object (Outlook)
keywords: vbaol11.chm3205
f1_keywords:
- vbaol11.chm3205
ms.prod: outlook
api_name:
- Outlook.ViewField
ms.assetid: 997319f0-7ff3-a712-8484-2e442965e187
ms.date: 06/08/2017
---


# ViewField Object (Outlook)

Represents a view field, used to display information in a view.


## Remarks

Use the  **[Add](viewfields-add-method-outlook.md)** method of the **[ViewFields](viewfields-object-outlook.md)** collection to add an Outlook item property to the following objects derived from the **[View](view-object-outlook.md)** object:


-  **[CardView](cardview-object-outlook.md)**
    
-  **[TableView](tableview-object-outlook.md)**
    
Use the  **[ColumnFormat](viewfield-columnformat-property-outlook.md)** property to access the **[ColumnFormat](columnformat-object-outlook.md)** object representing the display properties associated with the view field. Use the **[ViewXMLSchemaName](viewfield-viewxmlschemaname-property-outlook.md)** property to obtain the name of the view field as referenced in the XML definition of the view.


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
|[Application](viewfield-application-property-outlook.md)|
|[Class](viewfield-class-property-outlook.md)|
|[ColumnFormat](viewfield-columnformat-property-outlook.md)|
|[Parent](viewfield-parent-property-outlook.md)|
|[Session](viewfield-session-property-outlook.md)|
|[ViewXMLSchemaName](viewfield-viewxmlschemaname-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
