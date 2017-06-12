---
title: Explorer.ShowPane Method (Outlook)
keywords: vbaol11.chm2776
f1_keywords:
- vbaol11.chm2776
ms.prod: outlook
api_name:
- Outlook.Explorer.ShowPane
ms.assetid: 3d2c9dd5-b660-e160-36db-73c23f95a7a2
ms.date: 06/08/2017
---


# Explorer.ShowPane Method (Outlook)

Displays or hides a specific pane in the explorer.


## Syntax

 _expression_ . **ShowPane**( **_Pane_** , **_Visible_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pane_|Required| **[OlPane](olpane-enumeration-outlook.md)**|The pane to display.|
| _Visible_|Required| **Boolean**| **True** to make the pane visible, **False** to hide the pane.|

## Remarks




 **Note**  You can also use the  **[Visible](outlookbarpane-visible-property-outlook.md)** property of the **[OutlookBarPane](outlookbarpane-object-outlook.md)** object to display or hide the Outlook Bar.


## Example

This Microsoft Visual Basic for Applications (VBA) example uses the  **ShowPane** and **[IsPaneVisible](explorer-ispanevisible-method-outlook.md)** methods to hide the preview pane if it is visible or to display it if it is hidden.


```vb
Sub ShowHidePreviewPane() 
 
 Dim myOlExp As Outlook.Explorer 
 
 
 
 Set myOlExp = Application.ActiveExplorer 
 
 myOlExp.ShowPane olPreview, _ 
 
 Not myOlExp.IsPaneVisible(olPreview) 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

