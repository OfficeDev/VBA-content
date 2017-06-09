---
title: Using Events with ActiveX Controls
keywords: vbawd10.chm5214012
f1_keywords:
- vbawd10.chm5214012
ms.prod: word
ms.assetid: a344a964-e5f8-b6f1-e1eb-3c2d1fea9fb6
ms.date: 06/08/2017
---


# Using Events with ActiveX Controls

Word documents can contain ActiveX controls. Use the  **Control Toolbox** to insert ActiveX controls such as command buttons, check boxes, and list boxes. Use the following steps to add an ActiveX check box control with a **LostFocus** event.


1. On the Developer tab, click the  **Check Box** control. A check box control is inserted in the active document.
    
2. Right-click the check box control and click  **View Code**. Word switches to the Visual Basic Editor and displays the ThisDocument class module with the check box selected in the  **Object** drop-down list box.
    
3. Select the  **LostFocus** event from the **Procedure** drop-down list box. An empty procedure is added to the class module.
    
4. Add the Visual Basic instructions you want to run when the event occurs.
    

The following example shows a  **LostFocus** event procedure that runs when the focus is moved away from CheckBox1. The macro displays the state of CheckBox1 using the **Value** property ( **True** for selected and **False** for clear).




```vb
Private Sub CheckBox1_LostFocus() 
 MsgBox CheckBox1.Value 
End Sub
```

To see your event procedure run, switch back to Word with the document that includes the check box displayed. Click the  **Exit Design Mode** button on the **Control Toolbox**. Select or clear the check box and then click another element in the document. The check box control loses the focus and your LostFocus procedure runs; a message box is displayed with either "True" or "False."
Word implements the  **LostFocus** and **GotFocus** events for ActiveX controls in a Word document. The other events listed in the **Procedure** drop-down list box in are documented in Microsoft Forms Help.

