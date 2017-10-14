---
title: AutoFormatRule Object (Outlook)
keywords: vbaol11.chm3209
f1_keywords:
- vbaol11.chm3209
ms.prod: outlook
api_name:
- Outlook.AutoFormatRule
ms.assetid: 6d295c41-17f9-8e67-4595-4330fd3cec99
ms.date: 06/08/2017
---


# AutoFormatRule Object (Outlook)

Represents a formatting rule used by a  **[View](view-object-outlook.md)** object to determine how to format Outlook items displayed within that view.


## Remarks

Use the  **[Add](autoformatrules-add-method-outlook.md)** method or the **[Insert](autoformatrules-insert-method-outlook.md)** method of the **[AutoFormatRules](autoformatrules-object-outlook.md)** collection to create a new formatting rule for the following objects:


-  **[CalendarView](calendarview-object-outlook.md)**
    
-  **[CardView](cardview-object-outlook.md)**
    
-  **[TableView](tableview-object-outlook.md)**
    

### Built-In and Custom Formatting Rules

Microsoft Outlook provides a set of built-in formatting rules that can be disabled but cannot be removed or reordered. Custom formatting rules, defined either programmatically or by user action, cannot be moved above or between built-in formatting rules. Use the  **[Standard](autoformatrule-standard-property-outlook.md)** property to determine whether a formatting rule is built-in or custom.


### Applying Formatting Rules

Formatting rules are checked and applied against each Outlook item, in the order in which they are contained within the  **AutoFormatRules** collection. Use the **[Enabled](autoformatrule-enabled-property-outlook.md)** property to enable or disable a formatting rule, the **[Filter](autoformatrule-filter-property-outlook.md)** property to define the conditions an Outlook item must meet to be formatted by the formatting rule, and the **[Font](autoformatrule-font-property-outlook.md)** property to specify the format to be applied by the formatting rule.


## Example

The following Visual Basic for Applications (VBA) example enumerates the  **[AutoFormatRules](tableview-autoformatrules-property-outlook.md)** collection for the current **TableView** object, disabling any custom formatting rule contained by the collection.


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


## Properties



|**Name**|
|:-----|
|[Application](autoformatrule-application-property-outlook.md)|
|[Class](autoformatrule-class-property-outlook.md)|
|[Enabled](autoformatrule-enabled-property-outlook.md)|
|[Filter](autoformatrule-filter-property-outlook.md)|
|[Font](autoformatrule-font-property-outlook.md)|
|[Name](autoformatrule-name-property-outlook.md)|
|[Parent](autoformatrule-parent-property-outlook.md)|
|[Session](autoformatrule-session-property-outlook.md)|
|[Standard](autoformatrule-standard-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
