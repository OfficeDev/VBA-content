---
title: ObjectFrame.ObjectVerbsCount Property (Access)
keywords: vbaac10.chm11609
f1_keywords:
- vbaac10.chm11609
ms.prod: access
api_name:
- Access.ObjectFrame.ObjectVerbsCount
ms.assetid: 8c7a6302-cdf0-5997-7b71-65cfb6f0a7d3
ms.date: 06/08/2017
---


# ObjectFrame.ObjectVerbsCount Property (Access)

You can use the  **ObjectVerbsCount** property in Visual Basic to determine the number of verbs supported by an OLE object. Read-only **Long**.


## Syntax

 _expression_. **ObjectVerbsCount**

 _expression_ A variable that represents an **ObjectFrame** object.


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


[ObjectFrame Object](objectframe-object-access.md)

