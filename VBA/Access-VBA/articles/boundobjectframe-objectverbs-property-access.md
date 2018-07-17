---
title: BoundObjectFrame.ObjectVerbs Property (Access)
keywords: vbaac10.chm10900
f1_keywords:
- vbaac10.chm10900
ms.prod: access
api_name:
- Access.BoundObjectFrame.ObjectVerbs
ms.assetid: 892ace19-928e-aa58-4a71-6f38c64727ff
ms.date: 06/08/2017
---


# BoundObjectFrame.ObjectVerbs Property (Access)

You can use the  **ObjectVerbs** property in Visual Basic to determine the list of verbs an OLE object supports. Read-only **String**.


## Syntax

 _expression_. **ObjectVerbs**( ** _Index_** )

 _expression_ A variable that represents a **BoundObjectFrame** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|An element in the array of supported verbs. This is a zero-based index, meaning zero (0) represents the first verb in the array, one (1) represents the second verb in the array, and so on.|

## Remarks

This property setting isn't available in Design view.

You can use the  **ObjectVerbs** property with the **ObjectVerbsCount** property to display a list of the verbs supported by an OLE object. The **ObjectVerbs** property uses this list of verbs to determine which operation to perform when an OLE object is activated (when the **Action** property is set to **acOLEActivate** ).

The  **Verb** property setting is the position of a particular verb in the list of verbs returned by the **ObjectVerbs** property. For example, 1 specifies the first verb in the list (the Visual Basic command `ObjectVerbs(0)`, or the first verb in the  **ObjectVerbs** property array), 2 specifies the second verb in the list (the Visual Basic command `ObjectVerbs(1)`, or the second verb in the  **ObjectVerbs** property array), and so on.

The first verb in the  **ObjectVerbs** property array, called by the Visual Basic command `ObjectVerbs(0)`, is the default verb. If the  **Verb** property hasn't been set, this verb specifies the operation performed when the OLE object is activated.

The list of verbs an object supports varies, depending on the state of the object. To update the list of verbs an object supports, set the control's **Action** property to **acOLEFetchVerbs**. Be sure to update the list of verbs before presenting it to the user.


## Example

The following example returns the verbs supported by the OLE object in the OLE1 control and displays each verb in a message box.


```vb
Sub GetVerbList(frm As Form, OLE1 As Control) 
 Dim intX As Integer, intNumVerbs As Integer 
 Dim strVerbList As String 
 
 ' Update verb list. 
 With frm!OLE1 
 .Action = acOLEFetchVerbs 
 intNumVerbs = .ObjectVerbsCount 
 For intX = 0 To intNumVerbs - 1 
 strVerbList = strVerbList &; .ObjectVerbs(intX) &; "; " 
 Next intX 
 End With 
 
 ' Display verbs in message box. 
 MsgBox Left(strVerbList, Len(strVerbList) - 2) 
End Sub
```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

