---
title: COMAddIns Object (Office)
keywords: vbaof11.chm220000
f1_keywords:
- vbaof11.chm220000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.COMAddIns
ms.assetid: f6efa1cc-8d30-27d5-8b07-7ddad22f16ef
---


# COMAddIns Object (Office)

A collection of  **COMAddIn** objects that provide information about a COM add-in registered in the Windows registry.


## Example

Use the  **COMAddIns** property of the **Application** object to return the **COMAddIns** collection for a Microsoft Office host application. This collection contains all of the COM add-ins that are available to a given Office host application, and the **Count** property of the **COMAddins** collection returns the number of available COM add-ins, as in the following example.


```
MsgBox Application.COMAddIns.Count
```

Use the  **Update** method of the **COMAddins** collection to refresh the list of COM add-ins from the Windows registry, as in the following example.




```
Application.COMAddIns.Update
```

Use  **COMAddIns.Item(index)**, where _index_ is either an ordinal value that returns the COM add-in at that position in the **COMAddIns** collection, or a **String** value that represents the ProgID of the specified COM add-in. The following example displays a COM add-in's description text and ProgID (" **msodraa9.ShapeSelect** ") in a message box.




```
MsgBox Application.COMAddIns.Item("msodraa9.ShapeSelect").Description
```


## Methods



|**Name**|
|:-----|
|[Item](http://msdn.microsoft.com/library/bc9f4f41-fe52-1ba0-160c-0b1926194806%28Office.15%29.aspx)|
|[Update](http://msdn.microsoft.com/library/4cbaff64-10e8-d792-60b5-29f6de97dc8f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/d1ee6b80-0a48-33e8-3fc3-45bc73ad1413%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/5522bdc5-15b5-473f-94e3-5010a3d30f4a%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/dedee4b9-f340-d8fa-2285-3f32a1c4f00a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1d0adb7a-867f-0241-8f13-1ba3310f201b%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[COMAddIns Object Members](http://msdn.microsoft.com/library/0fc908fa-0846-07ca-d2a2-4c87525ae719%28Office.15%29.aspx)
