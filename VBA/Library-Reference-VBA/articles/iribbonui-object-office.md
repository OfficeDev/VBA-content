---
title: IRibbonUI Object (Office)
keywords: vbaof11.chm320000
f1_keywords:
- vbaof11.chm320000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.IRibbonUI
ms.assetid: d323aa21-de74-e821-c914-db71ef3b9c5e
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


## Methods



|**Name**|
|:-----|
|[ActivateTab](http://msdn.microsoft.com/library/32f5205c-6ab1-e3a6-6bae-5f36706c4d0d%28Office.15%29.aspx)|
|[ActivateTabMso](http://msdn.microsoft.com/library/74096b3b-c2a7-0247-f3a1-d5e5dc7286e1%28Office.15%29.aspx)|
|[ActivateTabQ](http://msdn.microsoft.com/library/bf664b52-2660-2ce7-a01b-83b459f66e09%28Office.15%29.aspx)|
|[Invalidate](http://msdn.microsoft.com/library/068cd459-76c2-b1d3-ed7d-50fa88c4db73%28Office.15%29.aspx)|
|[InvalidateControl](http://msdn.microsoft.com/library/33af7933-66f7-51e9-895e-07a6222973d2%28Office.15%29.aspx)|
|[InvalidateControlMso](http://msdn.microsoft.com/library/bfcca0e9-8696-6a0e-ff27-6dfde41dff93%28Office.15%29.aspx)|

## See also


#### Other resources


[IRibbonUI Object Members](http://msdn.microsoft.com/library/c6f6ec3b-3132-da29-ea08-70f20923d013%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
