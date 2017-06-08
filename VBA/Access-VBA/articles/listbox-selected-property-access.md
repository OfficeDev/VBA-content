---
title: ListBox.Selected Property (Access)
keywords: vbaac10.chm11207
f1_keywords:
- vbaac10.chm11207
ms.prod: access
api_name:
- Access.ListBox.Selected
ms.assetid: db30f166-c82b-2a77-6feb-bf03810fc36d
ms.date: 06/08/2017
---


# ListBox.Selected Property (Access)

You can use the  **Selected** property in Visual Basic to determine if an item in a list box is selected. Read/write **Long**.


## Syntax

 _expression_. **Selected**( ** _lRow_** )

 _expression_ A variable that represents a **ListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lRow_|Required|**Long**|The item in the list box. The first item is represented by a zero (0), the second by a one (1), and so on.|

## Remarks

The  **Selected** property is a zero-based array that contains the selected state of each item in a list box.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|The list box item is selected.|
|**False**|The list box item isn't selected.|
This property is available only at run time.

When a list box control's  **MultiSelect** property is set to None, only one item can have its **Selected** property set to **True**. When a list box control's **MultiSelect** property is set to Simple or Extended, any or all of the items can have their **Selected** property set to **True**. A multiple-selection list box bound to a field will always have a **Value** property equal to **Null**. You use the **Selected** property or the **ItemsSelected** collection to retrieve information about which items are selected.

You can use the  **Selected** property to select items in a list box by using Visual Basic. For example, the following expression selects the fifth item in the list:




```vb
Me!Listbox.Selected(4) = True
```


## Example

The following example uses the  **Selected** property to move selected items in the lstSource list box to the lstDestination list box. The lstDestination list box's **RowSourceType** property is set to Value List and the control's **RowSource** property is constructed from all the selected items in the lstSource control. The lstSource list box's **MultiSelect** property is set to Extended. The CopySelected( ) procedure is called from the cmdCopyItem command button.


```vb
Private Sub cmdCopyItem_Click() 
 CopySelected Me 
End Sub 
 
Public Sub CopySelected(ByRef frm As Form) 
 
 Dim ctlSource As Control 
 Dim ctlDest As Control 
 Dim strItems As String 
 Dim intCurrentRow As Integer 
 
 Set ctlSource = frm!lstSource 
 Set ctlDest = frm!lstDestination 
 
 For intCurrentRow = 0 To ctlSource.ListCount - 1 
 If ctlSource.Selected(intCurrentRow) Then 
 strItems = strItems &; ctlSource.Column(0, _ 
 intCurrentRow) &; ";" 
 End If 
 Next intCurrentRow 
 
 ' Reset destination control's RowSource property. 
 ctlDest.RowSource = "" 
 ctlDest.RowSource = strItems 
 
 Set ctlSource = Nothing 
 Set ctlDest = Nothing 
 
End Sub
```


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

