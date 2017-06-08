---
title: OrderField Object (Outlook)
keywords: vbaol11.chm3187
f1_keywords:
- vbaol11.chm3187
ms.prod: outlook
api_name:
- Outlook.OrderField
ms.assetid: 4ae32270-bde9-3178-bca3-f8d145779d3d
ms.date: 06/08/2017
---


# OrderField Object (Outlook)

Represents an order field, used to sort information in a view.


## Remarks

Use the  **[Add](viewfields-add-method-outlook.md)** method of the **[OrderFields](orderfields-object-outlook.md)** object to add an Outlook item property to the **SortFields** collection for the following objects derived from the **[View](view-object-outlook.md)** object:


-  **[BusinessCardView](businesscardview-object-outlook.md)**
    
-  **[CardView](cardview-object-outlook.md)**
    
-  **[IconView](iconview-object-outlook.md)**
    
-  **[TableView](tableview-object-outlook.md)**
    
Use the  **[ViewXMLSchemaName](orderfield-viewxmlschemaname-property-outlook.md)** property to obtain the name of the order field as referenced in the XML definition of the view.

 **OrderField** objects contained in an **OrderFields** collection are applied to Outlook items displayed in the view in the order in which the objects are contained in the collection. For each **OrderField** object, use the **[IsDescending](orderfield-isdescending-property-outlook.md)** property to determine whether to sort the contents of the order field in ascending or descending order.


## Example

The following Visual Basic for Applications (VBA) example iterates through the  **[SortFields](tableview-sortfields-property-outlook.md)** collection of the current **[TableView](tableview-object-outlook.md)** object, displaying the label and XML schema names of each **OrderField** object in the collection.


```
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
 
 strOutput = strOutput &amp; .ColumnFormat.Label &amp; _ 
 
 " (" &amp; .ViewXMLSchemaName &amp; ")" &amp; vbCrLf 
 
 End With 
 
 Next 
 
 
 
 ' Display a dialog box containing the concatenated 
 
 ' sort field information. 
 
 MsgBox strOutput 
 
 End If 
 
End Sub 
 

```


## Properties



|**Name**|
|:-----|
|[Application](orderfield-application-property-outlook.md)|
|[Class](orderfield-class-property-outlook.md)|
|[IsDescending](orderfield-isdescending-property-outlook.md)|
|[Parent](orderfield-parent-property-outlook.md)|
|[Session](orderfield-session-property-outlook.md)|
|[ViewXMLSchemaName](orderfield-viewxmlschemaname-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
