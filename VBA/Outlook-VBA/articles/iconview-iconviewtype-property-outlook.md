---
title: IconView.IconViewType Property (Outlook)
keywords: vbaol11.chm2576
f1_keywords:
- vbaol11.chm2576
ms.prod: outlook
api_name:
- Outlook.IconView.IconViewType
ms.assetid: 8255256d-eb71-6d3c-66bf-27aa5a103297
ms.date: 06/08/2017
---


# IconView.IconViewType Property (Outlook)

Returns or sets an  **[OlIconViewType](oliconviewtype-enumeration-outlook.md)** constant that determines how Outlook items are displayed in the **[IconView](iconview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **IconViewType**

 _expression_ A variable that represents an **IconView** object.


## Remarks

If the value of this property is set to  **olIconSortAndAutoArrange** , the value of the **[IconPlacement](iconview-iconplacement-property-outlook.md)** property is automatically set to **olIconSortAndAutoArrange** .


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

