---
title: Max, Min, Zoom Properties, Frame, ScrollBar Controls Example
keywords: fm20.chm5225158
f1_keywords:
- fm20.chm5225158
ms.prod: office
ms.assetid: 87bb60ba-4d1c-3160-b3d8-2e70019ec590
ms.date: 06/08/2017
---


# Max, Min, Zoom Properties, Frame, ScrollBar Controls Example

The following example uses the  **Zoom** property to shrink or enlarge the information displayed on a form, Page, or Frame. This example includes a **Frame**, a **TextBox** in the **Frame**, and a **ScrollBar**. The magnification level of the **Frame** changes through **Zoom**. The user can set **Zoom** by using the **ScrollBar**. The **TextBox** is present to demonstrate the effects of zooming.

This example also uses the  **Max** and **Min** properties to identify the range of acceptable values for the **ScrollBar**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:



- A  **Label** named Label1.
    
- A  **ScrollBar** named ScrollBar1.
    
- A second  **Label** named Label2.
    
- A  **Frame** named Frame1.
    
- A  **TextBox** named TextBox1 that is located inside Frame1.
    




```vb
Private Sub UserForm_Initialize() 
 ScrollBar1.Max = 400 
 ScrollBar1.Min = 10 
 ScrollBar1.Value = 100 
 
 Label1.Caption = "10 -----Percent of " _ 
 &; "Original Size---- 400" 
 Label2.Caption = ScrollBar1.Value 
 
 Frame1.TextBox1.Text = "Enter your text here." 
 Frame1.TextBox1.MultiLine = True 
 Frame1.TextBox1.WordWrap = True 
 
 Frame1.Zoom = ScrollBar1.Value 
End Sub 
 
Private Sub ScrollBar1_Change() 
 Frame1.Zoom = ScrollBar1.Value 
 Label2.Caption = ScrollBar1.Value 
End Sub
```


