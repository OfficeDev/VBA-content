---
title: ListBox Control, DataObject Object, MouseMove Event, StartDrag, SetText Methods Example
keywords: fm20.chm5225174
f1_keywords:
- fm20.chm5225174
ms.prod: office
ms.assetid: 83930d1d-a7e1-0c72-7e33-20922206c917
ms.date: 06/08/2017
---


# ListBox Control, DataObject Object, MouseMove Event, StartDrag, SetText Methods Example

The following example demonstrates a drag-and-drop operation from one  **ListBox** to another using a **DataObject** to contain the dragged text. This code sample uses the **SetText** and **StartDrag** methods in the **MouseMove** event to implement the drag-and-drop operation.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains two  **ListBox** controls named ListBox1 and ListBox2. You also need to add choices to the second **ListBox**.



```vb
Private Sub ListBox2_BeforeDragOver(ByVal Cancel As _ 
 MSForms.ReturnBoolean, ByVal Data As _ 
 MSForms.DataObject, ByVal X As Single, _ 
 ByVal Y As Single, ByVal DragState As Long, _ 
 ByVal Effect As MSForms.ReturnEffect, _ 
 ByVal Shift As Integer) 
 Cancel = True 
 Effect = 1 
End Sub 
 
Private Sub ListBox2_BeforeDropOrPaste(ByVal _ 
 Cancel As MSForms.ReturnBoolean, _ 
 ByVal Action As Long, ByVal Data As _ 
 MSForms.DataObject, ByVal X As Single, _ 
 ByVal Y As Single, ByVal Effect As _ 
 MSForms.ReturnEffect, ByVal Shift As Integer) 
 Cancel = True 
 Effect = 1 
 ListBox2.AddItem Data.GetText 
End Sub 
 
Private Sub ListBox1_MouseMove(ByVal Button As _ 
 Integer, ByVal Shift As Integer, ByVal X As _ 
 Single, ByVal Y As Single) 
 Dim MyDataObject As DataObject 
 If Button = 1 Then 
 Set MyDataObject = New DataObject 
 Dim Effect As Integer 
 MyDataObject.SetText ListBox1.Value 
 Effect = MyDataObject.StartDrag 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 For i = 1 To 10 
 ListBox1.AddItem "Choice " _ 
 &; (ListBox1.ListCount + 1) 
 Next i 
End Sub
```


