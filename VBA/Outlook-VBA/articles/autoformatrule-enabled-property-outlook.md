---
title: AutoFormatRule.Enabled Property (Outlook)
keywords: vbaol11.chm2709
f1_keywords:
- vbaol11.chm2709
ms.prod: outlook
api_name:
- Outlook.AutoFormatRule.Enabled
ms.assetid: b3a99916-83b8-68b8-5541-e4db7d0c9bb1
ms.date: 06/08/2017
---


# AutoFormatRule.Enabled Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether the formatting rule represented by the **[AutoFormatRule](autoformatrule-object-outlook.md)** object is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents an **AutoFormatRule** object.


## Example

The following Visual Basic for Applications (VBA) example enumerates the  **[AutoFormatRules](tableview-autoformatrules-property-outlook.md)** collection for the current **TableView** object, disabling any custom formatting rule contained by the collection.


```vb
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


## See also


#### Concepts


[AutoFormatRule Object](autoformatrule-object-outlook.md)

