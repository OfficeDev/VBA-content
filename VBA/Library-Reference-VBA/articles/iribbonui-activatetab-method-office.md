---
title: IRibbonUI.ActivateTab Method (Office)
keywords: vbaof11.chm320004
f1_keywords:
- vbaof11.chm320004
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.IRibbonUI.ActivateTab
ms.assetid: 32f5205c-6ab1-e3a6-6bae-5f36706c4d0d
---


# IRibbonUI.ActivateTab Method (Office)

Activates the specified custom tab. This method returns S_FALSE if there is no Ribbon or the Ribbon is collapsed.


## Syntax

 _expression_. **ActivateTab**( ** _ControlID_** )

 _expression_ An expression that returns a **IRibbonUI** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ControlID_|Required|**String**|Specifies the Id of the custom Ribbon tab to be activated.|

### Return Value

Nothing


## Example

The following code makes the custom tab the active tab.


```vb
Public myRibbon As IRibbonUI 
 
Sub tabActivate(ByVal control As IRibbonControl) 
 myRibbon.ActivateTab (control.ID) 
End Sub
```


## See also


#### Concepts


[IRibbonUI Object](iribbonui-object-office.md)

