---
title: AutoFormatRule.Filter Property (Outlook)
keywords: vbaol11.chm2708
f1_keywords:
- vbaol11.chm2708
ms.prod: outlook
api_name:
- Outlook.AutoFormatRule.Filter
ms.assetid: 9ae874ba-8d40-ac5d-42e3-8082d790ab3e
ms.date: 06/08/2017
---


# AutoFormatRule.Filter Property (Outlook)

Returns or sets a  **String** value that represents the filter for a custom formatting rule. Read/write.


## Syntax

 _expression_ . **Filter**

 _expression_ A variable that represents an **AutoFormatRule** object.


## Remarks

The value of this property is a DAV Searching and Locating (DASL) string that represents the current filter for the custom formatting rule. For more information about using DASL to filter items formatted by the formatting rule, see [Filtering Items](http://msdn.microsoft.com/library/4038e042-1b07-5d18-18b0-c2b58c9c42da%28Office.15%29.aspx). Setting this property to an empty string applies the custom formatting rule to all items displayed by the view.


 **Note**  This property returns an empty string for a standard formatting rule (an  **[AutoFormatRule](autoformatrule-object-outlook.md)** object with a **[Standard](autoformatrule-standard-property-outlook.md)** property value set to **True** ). An error occurs if you attempt to assign a value to this property for a standard formatting rule.


## Example

The following Visual Basic for Applications (VBA) example obtains a  **[View](view-object-outlook.md)** object by using the **[CurrentView](explorer-currentview-property-outlook.md)** property of the **[Explorer](explorer-object-outlook.md)** object, and then creates a new **AutoFormatRule** named "Handoff Messages." The **Filter** property of the **AutoFormatRule** object is set so that the formatting rule applies to any message in which the **[Subject](mailitem-subject-property-outlook.md)** property value starts with "HANDOFF". The sample then sets the properties of the **[Font](autoformatrule-font-property-outlook.md)** object for the **AutoFormatRule** object the so that messages to which the formatting rule applies are displayed in blue, bold, 8 point Courier New text.


```vb
Private Sub FormatHandoffMessages() 
 
 Dim objView As TableView 
 
 Dim objRule As AutoFormatRule 
 
 
 
 ' Check if the current view is a table view. 
 
 If Application.ActiveExplorer.CurrentView.ViewType = olTableView Then 
 
 
 
 ' Obtain a TableView object reference to the current view. 
 
 Set objView = Application.ActiveExplorer.CurrentView 
 
 
 
 ' Create a new rule that displays any message with a 
 
 ' subject line that starts with "HANDOFF" in 
 
 ' blue, bold, 8 point Courier New text. 
 
 Set objRule = objView.AutoFormatRules.Add("Handoff Messages") 
 
 With objRule 
 
 .Filter = """http://schemas.microsoft.com/mapi/proptag/0x0037001f""" &; _ 
 
 " CI_STARTSWITH 'HANDOFF'" 
 
 With .Font 
 
 .Name = "Courier New" 
 
 .Size = "8" 
 
 .Bold = True 
 
 .Color = olColorBlue 
 
 End With 
 
 End With 
 
 
 
 ' Save and apply the table view. 
 
 objView.Save 
 
 objView.Apply 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[AutoFormatRule Object](autoformatrule-object-outlook.md)

