---
title: IRibbonUI.Invalidate Method (Office)
keywords: vbaof11.chm320001
f1_keywords:
- vbaof11.chm320001
ms.prod: office
api_name:
- Office.IRibbonUI.Invalidate
ms.assetid: 068cd459-76c2-b1d3-ed7d-50fa88c4db73
ms.date: 06/08/2017
---


# IRibbonUI.Invalidate Method (Office)

Invalidates the cached values for all of the controls of the Ribbon user interface.


## Syntax

 _expression_. **Invalidate**

 _expression_ An expression that returns a **IRibbonUI** object.


## Remarks

You can customize the Ribbon UI by using callback procedures in COM add-ins. For each of the callbacks the add-in implements, the responses are cached. For example, if an add-in writer implements the  **getImage** callback procedure for a button, the function is called once, the image loads, and then if the image needs to be updated, the cached image is used instead of recalling the procedure. This process remains in-place until the add-in signals that the cached values are invalid by using the **Invalidate** method, at which time, the callback procedure is again called and the return response is cached. The add-in can then force an immediate update of the UI by calling the **Refresh** method.


## Example

In the following example, starting the host application triggers the  **onLoad** event procedure that then calls a procedure which creates an object representing the Ribbon UI. Next, a callback procedure is defined that invalidates all of the controls on the UI and then refreshes the UI.


```XML
<customUI … OnLoad="MyAddinInitialize" …>
```


```
Dim MyRibbon As IRibbonUI 
 
Sub MyAddInInitialize(Ribbon As IRibbonUI) 
 Set MyRibbon = Ribbon 
End Sub 
 
Sub myFunction() 
 MyRibbon.Invalidate() ' Invalidates the caches of all of this add-in's controls 
End Sub
```


## See also


#### Concepts


[IRibbonUI Object](iribbonui-object-office.md)
#### Other resources


[IRibbonUI Object Members](iribbonui-members-office.md)

