---
title: Pages.Remove Method (Access)
keywords: vbaac10.chm10129
f1_keywords:
- vbaac10.chm10129
ms.prod: access
api_name:
- Access.Pages.Remove
ms.assetid: 24dff544-d544-2be5-6506-66d3f1ab3a0f
ms.date: 06/08/2017
---


# Pages.Remove Method (Access)

The  **Remove** method removes a **[Page](page-object-access.md)** object from the **[Pages](pages-object-access.md)** collection of a tab control.


## Syntax

 _expression_. **Remove**( ** _Item_** )

 _expression_ A variable that represents a **Pages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Optional|**Variant**|An integer that specifies the index of the  **Page** object to be removed. The index of the **Page** object corresponds to the value of the **PageIndex** property for that **Page** object. If you omit this argument, the last **Page** object in the collection is removed.|

## Remarks

The  **Pages** collection is indexed beginning with zero. The leftmost page in the tab control has an index of 0, the page immediately to the right of the leftmost page has an index of 1, and so on.

You can remove a  **Page** object from the **Pages** collection of a tab control only when the form is in Design view.


## Example

The following example removes pages from a tab control:


```vb
Function RemovePage() As Boolean 
 Dim frm As Form 
 Dim tbc As TabControl, pge As Page 
 
 On Error GoTo Error_RemovePage 
 Set frm = Forms!Form1 
 Set tbc = frm!TabCtl0 
 tbc.Pages.Remove 
 RemovePage = True 
 
Exit_RemovePage: 
 Exit Function 
 
Error_RemovePage: 
 MsgBox Err &; ": " &; Err.Description 
 RemovePage = False 
 Resume Exit_RemovePage 
End Function
```


## See also


#### Concepts


[Pages Collection](pages-object-access.md)

