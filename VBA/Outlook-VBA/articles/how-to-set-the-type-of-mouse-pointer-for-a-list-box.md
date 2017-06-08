---
title: "How to: Set the Type of Mouse Pointer for a List Box"
keywords: olfm10.chm3077218
f1_keywords:
- olfm10.chm3077218
ms.prod: outlook
ms.assetid: 0db05edd-682f-cdc0-523e-c48e1a249017
ms.date: 06/08/2017
---


# How to: Set the Type of Mouse Pointer for a List Box

The following example demonstrates how to specify a mouse pointer that is appropriate for a specific control or situation. For the  **[ListBox](listbox-object-outlook-forms-script.md)** control, you can assign one of several available mouse pointers using the ** [ListBox.MousePointer](listbox-mousepointer-property-outlook-forms-script.md)** property.

This example works in the following ways:

- Choose a mouse pointer from the  **ListBox** to change the mouse pointer associated with the **ListBox**.
    
To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **ListBox** named ListBox1.
    



```vb
Dim ListBox1 
 
Sub Item_Open() 
 set ListBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("ListBox1") 
 'Load ListBox with MousePointer choices 
 ListBox1.Clear 
 ListBox1.AddItem "Default" 
 ListBox1.AddItem "Arrow" 
 ListBox1.AddItem "Cross" 
 ListBox1.AddItem "IBeam" 
 ListBox1.AddItem "SizeNESW" 
 ListBox1.AddItem "SizeNS" 
 ListBox1.AddItem "SizeNWSE" 
 ListBox1.AddItem "SizeWE" 
 ListBox1.AddItem "UpArrow" 
 ListBox1.AddItem "Hourglass" 
 ListBox1.AddItem "NoDrop" 
 ListBox1.AddItem "AppStarting" 
 ListBox1.AddItem "Help" 
 ListBox1.AddItem "SizeAll" 
End Sub 
 
Sub ListBox1_Click() 
 If IsNull(ListBox1.Value) = False Then 
 Select Case ListBox1.Value 
 Case "Default" 
 pointer = 0 'Standard pointer. 
 Case "Arrow" 
 pointer = 1 'Arrow. 
 Case "Cross" 
 pointer = 2 'Cross-hair pointer. 
 Case "IBeam" 
 pointer = 3 'I-beam. 
 Case "SizeNESW" 
 pointer = 6 'Double arrow pointing northeast and southwest. 
 Case "SizeNS" 
 pointer = 7 'Double arrow pointing north and south. 
 Case "SizeNWSE" 
 pointer = 8 'Double arrow pointing northwest and southeast. 
 Case "SizeWE" 
 pointer = 9 'Double arrow pointing west and east. 
 Case "UpArrow" 
 pointer = 10 'Up arrow. 
 Case "Hourglass" 
 pointer = 11 'Hourglass. 
 Case "NoDrop" 
 pointer = 12 '"Not" symbol (circle with a diagonal line) on top of the object being dragged. Indicates an invalid drop target. 
 Case "AppStarting" 
 pointer = 13 'Arrow with an hourglass. 
 Case "Help" 
 pointer = 14 'Arrow with a question mark. 
 Case "SizeAll" 
 pointer = 15 'Size all cursor (arrows pointing north, south, east, and west). 
 End Select 
 ListBox1.MousePointer = pointer 
 End If 
End Sub
```


