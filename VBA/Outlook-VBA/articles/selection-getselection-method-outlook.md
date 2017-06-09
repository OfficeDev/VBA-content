---
title: Selection.GetSelection Method (Outlook)
keywords: vbaol11.chm3533
f1_keywords:
- vbaol11.chm3533
ms.prod: outlook
api_name:
- Outlook.Selection.GetSelection
ms.assetid: c6af6665-d97d-3833-1014-5b43282bafc2
ms.date: 06/08/2017
---


# Selection.GetSelection Method (Outlook)

Returns a  **[Selection](selection-object-outlook.md)** object that contains the kind of objects specified by the _SelectionContents_ parameter, and that are currently selected in the active explorer.


## Syntax

 _expression_ . **GetSelection**( **_SelectionContents_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SelectionContents_|Required| **[OlSelectionContents](olselectioncontents-enumeration-outlook.md)**|Specifies the kind of objects in the selection to return.|

### Return Value

A  **Selection** object that contains the specified kind of objects that are selected in the active explorer.


## Remarks

Calling  **GetSelection** with **olConversationHeaders** as the argument returns a **Selection** object that has the **[Location](selection-location-property-outlook.md)** property equal to **OlSelectionLocation.olViewList** .

If the current view is not a conversation view, or, if  **Selection.Location** is not equal to **OlSelectionLocation.olViewList** , calling **GetSelection** with **olConversationHeaders** as the argument returns a **Selection** object with **[Selection.Count](selection-count-property-outlook.md)** equal to 0.


## See also


#### Concepts


[Selection Object](selection-object-outlook.md)
#### Other resources



[How to: Obtain and Enumerate Selected Conversations](http://msdn.microsoft.com/library/3bba1e98-b2eb-c53d-354a-bdd899b65a59%28Office.15%29.aspx)

