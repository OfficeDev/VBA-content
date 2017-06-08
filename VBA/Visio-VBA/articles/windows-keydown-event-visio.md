---
title: Windows.KeyDown Event (Visio)
keywords: vis_sdr.chm11751325
f1_keywords:
- vis_sdr.chm11751325
ms.prod: visio
api_name:
- Visio.Windows.KeyDown
ms.assetid: 3f8b0efc-071b-ddd9-7413-cff7b28a26ba
ms.date: 06/08/2017
---


# Windows.KeyDown Event (Visio)

Occurs when a keyboard key is pressed.


## Syntax

Private Sub  _expression_ _**KeyDown**( **_ByVal KeyCode As Long_** , **_ByVal KeyButtonState As Long_** , **_ByVal CancelDefault As Boolean_** )

 _expression_ A variable that represents a **Windows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The key that was pressed. See Remarks for possible values.|
| _KeyButtonState_|Required| **Long**|The state of the SHIFT and CTRL keys for the event. See Remarks for possible values.|
| _CancelDefault_|Required| **Boolean**| **False** if Microsoft Visio should process the message it receives from this event; otherwise, **True** .|

## Remarks

Possible values for  _KeyCode_ are declared in **KeyCodeConstants** in the Microsoft Visual Basic for Applications (VBA) library.

Possible values for  _KeyButtonState_ can be a combination of the values shown in the following table, which are declared in **VisKeyButtonFlags** in the Visio type library. For example, if _KeyButtonState_ returns 12, it indicates that the user held down both SHIFT and CTRL.



|**Constant **|**Value **|
|:-----|:-----|
| **visKeyControl **|8|
| **visKeyShift **|4|
| **visMouseLeft **|1|
| **visMouseMiddle **|16|
| **visMouseRight **|2|
If you set  _CancelDefault_ to **True** , Visio will not process the message received when the mouse button is clicked.

Unlike some other Visio events,  **KeyDown** does not have the prefix "Query," but it is nevertheless a query event. That is, you can cancel processing the message sent by **KeyDown** , either by setting _CancelDefault_ to **True** , or, if you are using the **VisEventProc** method to handle the event, by returning **True** . For more information, see the topics for the **VisEventProc** method and for any of the query events (for example, the **QueryCancelSuspend** event) in this Automation Reference.


 **Note**  Pressing an accelererator key combination, for example, CTRL + C, does not fire the  **KeyDown** event.

If you are using VBA, the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


## Example

This class module shows how to define a sink class called  **KeyboardListener** that listens for events fired by keyboard actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains event handlers for the **KeyDown** , **KeyPress** , and **KeyUp** events.

To run this example, insert a new class module in your VBA project, name it  **KeyboardListener** , and insert the following code in the module.




```vb
Dim WithEvents vsoWindow As Visio.Window 
 
Private Sub Class_Initialize() 
 
 Set vsoWindow = ActiveWindow 
 
End Sub 
 
Private Sub Class_Terminate() 
 
 Set vsoWindow = Nothing 
 
End Sub 
 
Private Sub vsoWindow_KeyDown(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyCode is "; KeyCode 
 Debug.Print "KeyButtonState is" ; KeyButtonState 
 
End Sub 
 
Private Sub vsoWindow_KeyPress(ByVal KeyAscii As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyAscii value is "; KeyAscii 
 
End Sub 
 
Private Sub vsoWindow_KeyUp(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyCode is "; KeyCode 
 Debug.Print "KeyButtonState is" ; KeyButtonState 
 
End Sub
```

Then, insert the following code in the  **ThisDocument** project.




```vb
Dim myKeyboardListener As KeyboardListener 
 
Private Sub Document_DocumentSaved(ByVal doc As IVDocument) 
 
 Set myKeyboardListener = New KeyboardListener 
 
End Sub 
 
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument) 
 
 Set myKeyboardListener = Nothing 
 
End Sub
```

Save the document to initialize the class, and then press any key to fire a  **KeyDown** event. In the **Immediate** window, the handler prints the code of the key that was pressed to fire the event and the state of the SHIFT and CTRL keys at the time the event fired.


