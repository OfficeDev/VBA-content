---
title: NavigationPane.DisplayedModuleCount Property (Outlook)
keywords: vbaol11.chm2792
f1_keywords:
- vbaol11.chm2792
ms.prod: outlook
api_name:
- Outlook.NavigationPane.DisplayedModuleCount
ms.assetid: f94018b1-95b9-403d-212b-e59e2bca9438
ms.date: 06/08/2017
---


# NavigationPane.DisplayedModuleCount Property (Outlook)

Returns or sets a  **Long** value that indicates the number of **[NavigationModule](navigationmodule-object-outlook.md)** objects displayed in the Navigation Pane. Read/write.


## Syntax

 _expression_ . **DisplayedModuleCount**

 _expression_ A variable that represents a **NavigationPane** object.


## Remarks

This property can only be set to a value between 0 and the value of the  **[Count](navigationmodules-count-property-outlook.md)** property for the **[Modules](navigationpane-modules-property-outlook.md)** collection of the **NavigationPane** object. If this property is set to a value greater than the maximum allowable value, the property value is automatically set to the maximum allowable value. An error occurs if this property is set to less than 0.

 If the **[IsCollapsed](navigationpane-iscollapsed-property-outlook.md)** property of the **[NavigationPane](navigationpane-object-outlook.md)** object is set to **False** , then this property value represents the number of navigation modules for which both icon and name are displayed in the Navigation Pane. If **IsCollapsed** is set to **True** , then the **DisplayedModuleCount** property value represents the number of navigation modules for which an icon is displayed in the Navigation Pane.

Setting the value of this property resizes the Modules section of the Navigation Pane to display more or fewer  **NavigationModule** objects as needed.


## Example

The following Visual Basic for Applications (VBA) example displays all navigation modules contained by the Navigation Pane, by setting the value of the  **DisplayedModuleCount** property equal to the **Count** property of the **Modules** collection for the **NavigationPane** object.


```vb
Sub DisplayAllModules() 
 
 Dim objPane As NavigationPane 
 
 
 
 ' Get the NavigationPane object for the 
 
 ' currently displayed Explorer object. 
 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
 
 
 ' Set the DisplayedModuleCount property to 
 
 ' display all modules contained by the 
 
 ' Navigation Pane. 
 
 objPane.DisplayedModuleCount = objPane.Modules.Count 
 
End Sub
```


## See also


#### Concepts


[NavigationPane Object](navigationpane-object-outlook.md)

