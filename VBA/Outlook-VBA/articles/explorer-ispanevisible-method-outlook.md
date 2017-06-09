---
title: Explorer.IsPaneVisible Method (Outlook)
keywords: vbaol11.chm2775
f1_keywords:
- vbaol11.chm2775
ms.prod: outlook
api_name:
- Outlook.Explorer.IsPaneVisible
ms.assetid: d547978a-f6b4-06ea-2358-8b6a81230240
ms.date: 06/08/2017
---


# Explorer.IsPaneVisible Method (Outlook)

Returns a  **Boolean** indicating whether a specific explorer pane is visible.


## Syntax

 _expression_ . **IsPaneVisible**( **_Pane_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pane_|Required| **[OlPane](olpane-enumeration-outlook.md)**|The pane to check.|

### Return Value

 **True** if the specified pane is displayed in the explorer; otherwise, **False** .


## Remarks

You can also use the  **[Visible](outlookbarpane-visible-property-outlook.md)** property of the **[OutlookBarPane](outlookbarpane-object-outlook.md)** object to determine whether the **Shortcuts** pane is visible.


## Example

This Microsoft Visual Basic for Applications (VBA) sample uses the  **IsPaneVisible** method to determine whether the preview pane is visible and uses the **[ShowPane](explorer-showpane-method-outlook.md)** method to display it if it is not visible. Use the **olNavigationPane** constant to hide or display the Navigation Pane.


```vb
Sub HidePreviewPane() 
 
 Dim myOlExp As Outlook.Explorer 
 
 Set myOlExp = Application.ActiveExplorer 
 
 If myOlExp.IsPaneVisible(olPreview) = False Then 
 
 myOlExp.ShowPane olPreview, True 
 
 End If 
 
 Set myOlExp = Nothing 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

