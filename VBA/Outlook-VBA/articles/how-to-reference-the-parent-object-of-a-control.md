---
title: "How to: Reference the Parent Object of a Control"
keywords: olfm10.chm3077227
f1_keywords:
- olfm10.chm3077227
ms.prod: outlook
ms.assetid: b870fcfb-0ff9-ad87-985e-61ef1362d449
ms.date: 06/08/2017
---


# How to: Reference the Parent Object of a Control

The following example uses the Microsoft Forms 2.0  **Parent** property to refer to the control, form, or other object that contains a specific control or object.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- Two  **[Label](label-object-outlook-forms-script.md)** controls named Label1 and Label2.
    
- A  **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1.
    
- One or more additional controls of your choice.
    



```vb
Dim MyControl 
Dim MyParent 
Dim ControlsIndex 
 
Sub Item_Open() 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton1 
 ControlsIndex = 0 
 CommandButton1.Caption = "Get Control and Parent" 
 CommandButton1.AutoSize = True 
 CommandButton1.WordWrap = True 
End Sub 
 
Sub CommandButton1_Click() 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Label1 
 Set Label2 = Item.GetInspector.ModifiedFormPages("P.2").Label2 
 
 'Process Controls collection for UserForm 
 Set MyControl = Item.GetInspector.ModifiedFormPages("P.2").Controls.Item(ControlsIndex) 
 Set MyParent = MyControl.Parent 
 Label1.Caption = MyControl.Name 
 Label2.Caption = MyParent.Name 
 
 'Prepare index for next control on Userform 
 ControlsIndex = ControlsIndex + 1 
 If ControlsIndex >= Item.GetInspector.ModifiedFormPages("P.2").Controls.Count Then 
 ControlsIndex = 0 
 End If 
End Sub
```


