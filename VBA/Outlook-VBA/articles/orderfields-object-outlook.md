---
title: OrderFields Object (Outlook)
keywords: vbaol11.chm3186
f1_keywords:
- vbaol11.chm3186
ms.prod: outlook
api_name:
- Outlook.OrderFields
ms.assetid: e115fb80-352d-fd2e-c1c3-d266776fe122
ms.date: 06/08/2017
---


# OrderFields Object (Outlook)

Represents the collection of  **[OrderField](orderfield-object-outlook.md)** objects in a view.


## Remarks

The  **OrderFields** collection represents the Outlook item properties used to sort Outlook items displayed in the view. Use the **[Add](orderfields-add-method-outlook.md)** method or the **OrderFields** collection to create a new order field for the following objects derived from the **[View](view-object-outlook.md)** object:


-  **[BusinessCardView](businesscardview-object-outlook.md)**
    
-  **[CardView](cardview-object-outlook.md)**
    
-  **[IconView](iconview-object-outlook.md)**
    
-  **[PeopleView](peopleview-object-outlook.md)**
    
-  **[TableView](tableview-object-outlook.md)**
    
 **OrderField** objects contained in an **OrderFields** collection are applied to Outlook items displayed in the view in the order in which the objects are contained in the collection.


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


## Methods



|**Name**|
|:-----|
|[Add](orderfields-add-method-outlook.md)|
|[Insert](orderfields-insert-method-outlook.md)|
|[Item](orderfields-item-method-outlook.md)|
|[Remove](orderfields-remove-method-outlook.md)|
|[RemoveAll](orderfields-removeall-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](orderfields-application-property-outlook.md)|
|[Class](orderfields-class-property-outlook.md)|
|[Count](orderfields-count-property-outlook.md)|
|[Parent](orderfields-parent-property-outlook.md)|
|[Session](orderfields-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
