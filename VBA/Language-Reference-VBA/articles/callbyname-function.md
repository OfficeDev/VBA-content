---
title: CallByName Function
keywords: vblr6.chm1020905
f1_keywords:
- vblr6.chm1020905
ms.prod: office
ms.assetid: e76dece5-244f-9514-4ccf-d993d6476061
ms.date: 06/08/2017
---


# CallByName Function



Executes a method of an object, or sets or returns a property of an [object](vbe-glossary.md).
 **Syntax**
 **CallByName( _object_**_, procname, calltype,[args()]_**)**
The  **CallByName** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_object_**|Required;  **Variant** ( **Object** ). The name of the object on which the function will be executed.|
|**_procname_**|Required;  **Variant** ( **String** ). A string expression containing the name of a property or method of the object.|
|**_calltype_**|Required;  **Constant**. A constant of type **vbCallType** representing the type of procedure being called.|
| _args()_|Optional:  **Variant** ( **Array** ).|
 **Remarks**
The  **CallByName** function is used to get or set a property, or invoke a method at run time using a string name.
In the following example, the first line uses  **CallByName** to set the **MousePointer** property of a text box, the second line gets the value of the **MousePointer** property, and the third line invokes the **Move** method to move the text box:



```
CallByName Text1, "MousePointer", vbLet, vbCrosshair
Result = CallByName (Text1, "MousePointer", vbGet)
CallByName Text1, "Move", vbMethod, 100, 100
```


## Example

This example uses the  **CallByName** function to invoke the **Move** method of a Command button.

The example also uses a form ( `Form1`) with a button ( `Command1`), and a label ( `Label1`). When the form is loaded, the  **Caption** property of the label is set to "Move", the name of the method to invoke. When you click the button, the **CallByName** function invokes the method to change the location of the button.




```vb
Option Explicit

Private Sub Form_Load()
    Label1.Caption = "Move"        ' Name of Move method.
End Sub

Private Sub Command1_Click()
    If Command1.Left <> 0 Then
        CallByName Command1, Label1.Caption, vbMethod, 0, 0
    Else
        CallByName Command1, Label1.Caption, vbMethod, 500, 500
    End If
```


