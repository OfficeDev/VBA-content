---
title: Pages.Add Method (Access)
keywords: vbaac10.chm10128
f1_keywords:
- vbaac10.chm10128
ms.prod: access
api_name:
- Access.Pages.Add
ms.assetid: f7235fb2-d775-85ea-7c50-62fa3f663d32
ms.date: 06/08/2017
---


# Pages.Add Method (Access)

The  **Add** method adds a new **[Page](page-object-access.md)** object to the **[Pages](pages-object-access.md)** collection of a tab control.


## Syntax

 _expression_. **Add**( ** _Before_** )

 _expression_ A variable that represents a **Pages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Optional|**Variant**|An  **Integer** that specifies the index of the **Page** object before which the new **Page** object should be added. The index of the **Page** object corresponds to the value of the **PageIndex** property for that **Page** object. If you omit this argument, the new **Page** object is added to the end of the collection.|

### Return Value

Page


## Remarks

The first  **Page** object in the **Pages** collection corresponds to the leftmost page in the tab control and has an index of 0. The second **Page** object is immediately to the right of the first page and has an index of 1, and so on for all the **Page** objects in the tab control.

If you specify 0 for the  _Before_ argument, the new **Page** object is added before the first **Page** object in the **Pages** collection. The new **Page** object then becomes the first **Page** object in the collection, with an index of 0.

You can add a  **Page** object to the **Pages** collection of a tab control only when the form is in Design view.


## Example

The following example adds a page to a tab control on a form that's in Design view. To try this example, create a new form named Form1 with a tab control named TabCtl0. Paste the following code into a standard module and run it:


```vb
Function AddPage() As Boolean 
 Dim frm As Form 
 Dim tbc As TabControl, pge As Page 
 
 On Error GoTo Error_AddPage 
 Set frm = Forms!Form1 
 Set tbc = frm!TabCtl0 
 tbc.Pages.Add 
 AddPage = True 
 
Exit_AddPage: 
 Exit Function 
 
Error_AddPage: 
 MsgBox Err &; ": " &; Err.Description 
 AddPage = False 
 Resume Exit_AddPage 
End Function
```


## See also


#### Concepts


[Pages Collection](pages-object-access.md)

