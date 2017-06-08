---
title: BoundObjectFrame.ObjectVerbsCount Property (Access)
keywords: vbaac10.chm10955
f1_keywords:
- vbaac10.chm10955
ms.prod: access
api_name:
- Access.BoundObjectFrame.ObjectVerbsCount
ms.assetid: 518eff16-aef0-9e3e-2e03-af036117a152
ms.date: 06/08/2017
---


# BoundObjectFrame.ObjectVerbsCount Property (Access)

You can use the  **ObjectVerbsCount** property in Visual Basic to determine the number of verbs supported by an OLE object. Read-only **Long**.


## Syntax

 _expression_. **ObjectVerbsCount**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **ObjectVerbsCount** property setting is a value that specifies the number of elements in the **ObjectVerbs** property array.

This property setting isn't available in Design view.

The list of verbs an OLE object supports may vary, depending on the state of the object. To update the list of supported verbs, set the control's **Action** property to **acOLEFetchVerbs**.


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

