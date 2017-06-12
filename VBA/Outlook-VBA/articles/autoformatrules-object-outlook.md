---
title: AutoFormatRules Object (Outlook)
keywords: vbaol11.chm3210
f1_keywords:
- vbaol11.chm3210
ms.prod: outlook
api_name:
- Outlook.AutoFormatRules
ms.assetid: 74514b71-964c-f17b-4df6-e1a5c5ed2b52
ms.date: 06/08/2017
---


# AutoFormatRules Object (Outlook)

Represents the collection of  **[AutoFormatRule](autoformatrule-object-outlook.md)** objects in a view.


## Remarks

Use the  **[Add](autoformatrules-add-method-outlook.md)** method or the **[Insert](autoformatrules-insert-method-outlook.md)** method of the **AutoFormatRules** collection to create a new formatting rule for the following objects derived from the **[View](view-object-outlook.md)** object:


-  **[BusinessCardView](businesscardview-object-outlook.md)**
    
-  **[CalendarView](calendarview-object-outlook.md)**
    
-  **[CardView](cardview-object-outlook.md)**
    
-  **[IconView](iconview-object-outlook.md)**
    
-  **[TableView](tableview-object-outlook.md)**
    
-  **[TimelineView Object](timelineview-object-outlook.md)**
    
 **AutoFormatRule** objects contained in an **AutoFormatRules** collection are applied to each Outlook item in the order in which they are contained in the collection. Changes to **AutoFormatRule** objects are persisted only if the **[Save](autoformatrules-save-method-outlook.md)** method of the **AutoFormatRules** collection is called.


## Example

The following Visual Basic for Applications (VBA) example enumerates the  **AutoFormatRules** collection for the current **TableView** object, disabling any custom formatting rule contained by the collection.


```
Private Sub DisableCustomAutoFormatRules() 
 
 Dim objTableView As TableView 
 
 Dim objRule As AutoFormatRule 
 
 
 
 ' Check if the current view is a table view. 
 
 If Application.ActiveExplorer.CurrentView.ViewType = olTableView Then 
 
 
 
 ' Obtain a TableView object reference to the current view. 
 
 Set objView = Application.ActiveExplorer.CurrentView 
 
 
 
 ' Enumerate the AutoFormatRules collection for 
 
 ' the table view, disabling any custom formatting 
 
 ' rule defined for the view. 
 
 For Each objRule In objView.AutoFormatRules 
 
 If Not objRule.Standard Then 
 
 objRule.Enabled = False 
 
 End If 
 
 Next 
 
 
 
 ' Save and apply the table view. 
 
 objView.Save 
 
 objView.Apply 
 
 End If 
 
End Sub 
 

```


## Methods



|**Name**|
|:-----|
|[Add](autoformatrules-add-method-outlook.md)|
|[Insert](autoformatrules-insert-method-outlook.md)|
|[Item](autoformatrules-item-method-outlook.md)|
|[Remove](autoformatrules-remove-method-outlook.md)|
|[RemoveAll](autoformatrules-removeall-method-outlook.md)|
|[Save](autoformatrules-save-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](autoformatrules-application-property-outlook.md)|
|[Class](autoformatrules-class-property-outlook.md)|
|[Count](autoformatrules-count-property-outlook.md)|
|[Parent](autoformatrules-parent-property-outlook.md)|
|[Session](autoformatrules-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
