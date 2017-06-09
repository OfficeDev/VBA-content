---
title: IconView.IconPlacement Property (Outlook)
keywords: vbaol11.chm2577
f1_keywords:
- vbaol11.chm2577
ms.prod: outlook
api_name:
- Outlook.IconView.IconPlacement
ms.assetid: 8b6f403e-1eeb-6f23-b780-808d5f5b5bc6
ms.date: 06/08/2017
---


# IconView.IconPlacement Property (Outlook)

Returns or sets an  **[OlIconViewPlacement](oliconviewplacement-enumeration-outlook.md)** value that determines how icons are placed within the **[IconView](iconview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **IconPlacement**

 _expression_ A variable that represents an **IconView** object.


## Remarks

If the value of the  **[IconViewType](iconview-iconviewtype-property-outlook.md)** property is set to **olIconViewList** , the value of this property is automatically set to **olIconSortAndAutoArrange** .


## Example

The following Visual Basic for Applications (VBA) example configures the current  **IconView** object to display Outlook items as a sorted, auto-arranged set of large icons.


```vb
Sub ConfigureIconView() 
 Dim objIconView As IconView 
 
 ' Check if the current view is an icon view. 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 olIconView Then 
 
 ' Obtain a IconView object reference for the 
 ' current icon view. 
 Set objIconView = _ 
 Application.ActiveExplorer.CurrentView 
 
 With objIconView 
 ' Display items in the icon view as a 
 ' set of large icons. 
 .IconViewType = olIconLarge 
 
 ' Sort and auto arrange the items 
 ' within the icon view. 
 .IconPlacement = olIconSortAndAutoArrange 
 
 ' Save the icon view. 
 .Save 
 End With 
 End If 
 
End Sub
```


## See also


#### Concepts


[IconView Object](iconview-object-outlook.md)

