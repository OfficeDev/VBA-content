---
title: MouseIcon, MousePointer Properties Example
keywords: fm20.chm5225143
f1_keywords:
- fm20.chm5225143
ms.prod: office
ms.assetid: 8abdcd9b-3199-4e06-490f-3f945d8f6013
ms.date: 06/08/2017
---


# MouseIcon, MousePointer Properties Example

The following example demonstrates how to specify a mouse pointer that is appropriate for a specific control or situation. You can assign one of several available mouse pointers using the  **MousePointer** property; or, you can assign a custom icon using the **MousePointer** and **MouseIcon** properties.

This example works in the following ways:




- Choose a mouse pointer from the  **ListBox** to change the mouse pointer associated with the first **CommandButton**.
    
- Click the first  **CommandButton** to associate its mouse pointer with the second **CommandButton**.
    
- Click the second  **CommandButton** to load a custom icon for its mouse pointer.
    

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:


- Two  **CommandButton** controls named CommandButton1 and CommandButton2.
    
- A  **ListBox** named ListBox1.
    


 **Note**  This example uses two icon files (identified by the .ico file extention in Windows) that are loaded using the  **LoadPicture** function. You should edit each **LoadPicture** function call to specify an icon file that resides on your system.




```vb
Private Sub ListBox1_Click() 
 If IsNull(ListBox1.Value) = False Then 
 CommandButton1.MousePointer = ListBox1.Value 
 
 If CommandButton1.MousePointer = _ 
 fmMousePointerCustom Then 
 CommandButton1.MouseIcon = _ 
 LoadPicture("c:\msvc20\cdk32\" _ 
 &; "samples\circ1\bix.ico") 
 End If 
 End If 
End Sub 
 
Private Sub CommandButton1_Click() 
 CommandButton2.MousePointer = CommandButton1.MousePointer 
 
 If CommandButton2.MousePointer = fmMousePointerCustom Then 
 CommandButton2.MouseIcon = CommandButton1.MouseIcon 
 End If 
End Sub 
 
Private Sub CommandButton2_Click() 
 CommandButton2.MousePointer = fmMousePointerCustom 
 CommandButton2.MouseIcon = LoadPicture("c:\msvc20\cdk32\samples\push\push.ico") 
End Sub 
 
Private Sub UserForm_Initialize() 
 'Load ListBox with MousePointer choices 
 ListBox1.ColumnCount = 2 
 
 ListBox1.AddItem "fmMousePointerDefault" 
 ListBox1.List(0, 1) = fmMousePointerDefault 
 ListBox1.AddItem "fmMousePointerArrow" 
 ListBox1.List(1, 1) = fmMousePointerArrow 
 ListBox1.AddItem "fmMousePointerCross" 
 ListBox1.List(2, 1) = fmMousePointerCross 
 
 ListBox1.AddItem "fmMousePointerIBeam" 
 ListBox1.List(3, 1) = fmMousePointerIBeam 
 ListBox1.AddItem "fmMousePointerSizeNESW" 
 ListBox1.List(4, 1) = fmMousePointerSizeNESW 
 ListBox1.AddItem "fmMousePointerSizeNS" 
 ListBox1.List(5, 1) = fmMousePointerSizeNS 
 
 ListBox1.AddItem "fmMousePointerSizeNWSE" 
 ListBox1.List(6, 1) = fmMousePointerSizeNWSE 
 ListBox1.AddItem "fmMousePointerSizeWE" 
 ListBox1.List(7, 1) = fmMousePointerSizeWE 
 ListBox1.AddItem "fmMousePointerUpArrow" 
 ListBox1.List(8, 1) = fmMousePointerUpArrow 
 
 ListBox1.AddItem "fmMousePointerHourglass" 
 ListBox1.List(9, 1) = fmMousePointerHourGlass 
 ListBox1.AddItem "fmMousePointerNoDrop" 
 ListBox1.List(10, 1) = fmMousePointerNoDrop 
 ListBox1.AddItem "fmMousePointerAppStarting" 
 ListBox1.List(11, 1) = fmMousePointerAppStarting 
 
 ListBox1.AddItem "fmMousePointerHelp" 
 ListBox1.List(12, 1) = fmMousePointerHelp 
 ListBox1.AddItem "fmMousePointerSizeAll" 
 ListBox1.List(13, 1) = fmMousePointerSizeAll 
 ListBox1.AddItem "fmMousePointerCustom" 
 ListBox1.List(14, 1) = fmMousePointerCustom 
 
 ListBox1.BoundColumn = 2 
 ListBox1.Value = fmMousePointerDefault 
 
 MsgBox "ListBox1.Value =" &; ListBox1.Value &; "." 
 CommandButton1.MousePointer = ListBox1.Value 
End Sub
```


