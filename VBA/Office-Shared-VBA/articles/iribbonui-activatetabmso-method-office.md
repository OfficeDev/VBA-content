---
title: IRibbonUI.ActivateTabMso Method (Office)
keywords: vbaof11.chm320005
f1_keywords:
- vbaof11.chm320005
ms.prod: office
api_name:
- Office.IRibbonUI.ActivateTabMso
ms.assetid: 74096b3b-c2a7-0247-f3a1-d5e5dc7286e1
ms.date: 06/08/2017
---


# IRibbonUI.ActivateTabMso Method (Office)

Activates the specified built-in tab.


## Syntax

 _expression_. **ActivateTabMso**( **_ControlID_** )

 _expression_ An expression that returns a **IRibbonUI** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ControlID_|Required|**String**|Specifies the Id of the custom Ribbon tab to be activated.|

### Return Value

Nothing


## Example

The following code make a built-in tab as specified by the control ID the active tab.


```
Public myRibbon As IRibbonUI 
 
Sub tabActivate(ByVal control As IRibbonControl) 
 myRibbon.ActivateTabMso (control.ID) 
End Sub
```


## See also


#### Concepts


[IRibbonUI Object](iribbonui-object-office.md)
#### Other resources


[IRibbonUI Object Members](iribbonui-members-office.md)

