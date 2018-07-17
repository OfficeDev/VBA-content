---
title: IRibbonUI Object (Office)
keywords: vbaof11.chm320000
f1_keywords:
- vbaof11.chm320000
ms.prod: office
api_name:
- Office.IRibbonUI
ms.assetid: d323aa21-de74-e821-c914-db71ef3b9c5e
ms.date: 06/08/2017
---


# IRibbonUI Object (Office)

The object that is returned by the  **onLoad** procedure specified on the **customUI** tag. The object contains methods for invalidating control properties and for refreshing the user interface.


## Remarks

You can customize the Ribbon user interface (UI) by using callback procedures in COM add-ins. When the host application starts, the  **onLoad** callback procedure is called. The callback procedure then returns a **IRibbonUI** object pointing to the user interface (UI). YOu can use that object to invoke the **Invalidate**, **InvalidateControl**, and **Refresh** methods.


## Example

In the following example, starting the host application triggers the  **onLoad** event procedure that then calls a procedure which creates a **IRibbonUI** object representing the Ribbon UI. Next, a callback procedure is defined that invalidates all of the cached controls and then refreshes the UI.


```XML
<customUI … OnLoad="MyAddInInitialize" …>
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


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[IRibbonUI Object Members](iribbonui-members-office.md)

