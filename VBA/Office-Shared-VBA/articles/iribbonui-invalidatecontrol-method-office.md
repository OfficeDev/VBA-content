---
title: IRibbonUI.InvalidateControl Method (Office)
keywords: vbaof11.chm320002
f1_keywords:
- vbaof11.chm320002
ms.prod: office
api_name:
- Office.IRibbonUI.InvalidateControl
ms.assetid: 33af7933-66f7-51e9-895e-07a6222973d2
ms.date: 06/08/2017
---


# IRibbonUI.InvalidateControl Method (Office)

Invalidates the cached value for a single control on the Ribbon user interface.


## Syntax

 _expression_. **InvalidateControl**( **_bstrControlID_** )

 _expression_ An expression that returns a **IRibbonUI** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrControlID_|Required|**String**|Specifies the ID of the control that will be invalidated.|

## Remarks

You can customize the Ribbon UI by using callback procedures in COM add-ins. For each of the callbacks the add-in implements, the responses are cached. For example, if an add-in writer implements the  **getImage** callback procedure for a button, the function is called once, the image loads, and then if the image needs to be updated, the cached image is used instead of recalling the procedure. This process remains in-place for the control until the add-in signals that the cached values are invalid by using the **InvalidateControl** method, at which time, the callback procedure is again called and the return response is cached.


## Example

In the following example, starting the host application triggers the  **onLoad** event procedure that then calls a procedure which creates an object representing the Ribbon UI. Next, a callback procedure is defined that invalidates a control on the UI and then refreshes the UI.


```XML
<customUI … OnLoad="MyAddInInitialize" …>
```


```
Dim MyRibbon As IRibbonUI 
 
Sub MyAddInInitialize(Ribbon As IRibbonUI) 
 Set MyRibbon = Ribbon 
End Sub 
 
Sub myFunction() 
 MyRibbon.InvalidateControl("control1") ' Invalidates the cache of a single control 
End Sub
```


## See also


#### Concepts


[IRibbonUI Object](iribbonui-object-office.md)
#### Other resources


[IRibbonUI Object Members](iribbonui-members-office.md)

