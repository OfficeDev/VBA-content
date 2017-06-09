---
title: Form.Filter Event (Access)
keywords: vbaac10.chm13660
f1_keywords:
- vbaac10.chm13660
ms.prod: access
api_name:
- Access.Form.Filter
ms.assetid: 265f3397-3dc9-21b3-ebac-55fb4e1261c0
ms.date: 06/08/2017
---


# Form.Filter Event (Access)

Occurs when the user opens a filter window by clicking  **Filter by Form**,  **Advanced Filter/Sort**, or  **Server Filter By Form**.


## Syntax

 _expression_. **Filter**( ** _Cancel_**, ** _FilterType_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines whether to open the filter window. Setting the Cancel argument to  **True** (?1) prevents the filter window from opening. You can also use the **CancelEvent** method of the **DoCmd** object to cancel opening the filter window.|
| _FilterType_|Required|**Integer**|The filter window the user is trying to open. The FilterType argument can be one of the following intrinsic constants:<ul><li><b>acFilterByForm</b></li><li><b>acFilterAdvanced</b></li><li><b>acServerFilterByForm</b></li></ul>|


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnFilter** property to the name of the macro or to [Event Procedure].

You can use the  **Filter** event to:


- Remove any previous filter for the form. To do this, set the  **Filter** or **ServerFilter** property of the form to a zero-length string (" ") in the **Filter** macro or event procedure. This is especially useful if you want to make sure extraneous criteria don't appear in the new filter. For example, when you use the Filter By Selection feature, the criteria you use (the selected text in the form) is added to the **Filter** or **ServerFilter** property WHERE clause expression, and appears in both the **Filter By Form** window and the **Advanced Filter/Sort** window or the **Server Filter By Form** window. You can remove these old criteria by using the **Filter** event.
    
- Enter default settings for the new filter. To do this, set the  **Filter** or **ServerFilter** property to include these criteria. For example, you may want all filters for a Products form to display only current products (products for which the Discontinued control in the Products form isn't selected).
    
- Use your own custom filter window instead of one of the Microsoft Access filter windows. When the  **Filter** event occurs, you can open your own custom form and use the entries on this form to set the **Filter** or **ServerFilter** property and filter the original form. When the user closes this custom form, set the **FilterOn** or **ServerFilterByForm** property of the original form to **True** (?1) to apply the filter. Canceling the **Filter** event prevents the Microsoft Access filter window from opening.
    
- Prevent certain controls on the form from appearing or being used in the  **Filter By Form** or **Server Filter By Form** window. If you hide or disable a control in the **Filter** macro or event procedure, the control is hidden or disabled in the **Filter By Form** or **Server Filter By Form** window, and can't be used to set filter criteria. You can then use the **ApplyFilter** event to show or enable this control after the filter is applied, or when the filter is removed from the form.
    

## Example

The following example shows how to disable the TotalDue control on an Orders form when the user tries to create a filter, so that the user can't filter on this field. Any records that have a TotalDue value and meet the other filter criteria will always be shown on the filtered form. This example also forces the user to use the  **Filter By Form** window and not the **Advanced Filter/Sort** window.

To try this example, add the following event procedure to an Orders form that contains a TotalDue control. Try to create a filter by using the  **Advanced Filter/Sort** window that uses the TotalDue control. Also try creating the same filter by using the **Filter By Form** window.




```vb
Private Sub Form_Filter(Cancel As Integer, FilterType As Integer) 
    If FilterType = acFilterByForm Then 
        Forms!Orders!TotalDue.Enabled = False 
    ElseIf FilterType = acFilterAdvanced Then 
        MsgBox "The best way to filter this form is to use the " _ 
            &; "Filter By Form command or toolbar button.", vbOKOnly + vbInformation 
        Cancel = True 
    End If 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

