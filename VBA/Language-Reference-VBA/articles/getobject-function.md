---
title: GetObject Function
keywords: vblr6.chm1010959
f1_keywords:
- vblr6.chm1010959
ms.prod: office
ms.assetid: 6c313a4c-dac9-9115-95db-3fde52a5e888
ms.date: 06/08/2017
---


# GetObject Function



Returns a reference to an object provided by an ActiveX component.
 **Syntax**
 **GetObject(** [ **_pathname_** ] [ **,  _class_** ] **)**
The  **GetObject** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_pathname_**|Optional;  **Variant** ( **String** ). The full path and name of the file containing the object to retrieve. If **_pathname_** is omitted, **_class_** is required.|
|**_class_**|Optional;  **Variant** ( **String** ). A string representing the[class](vbe-glossary.md) of the object.|
The  **_class_**[argument](vbe-glossary.md) uses the syntax _appname_**.**_objecttype_ and has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _appname_|Required;  **Variant** ( **String** ). The name of the application providing the object.|
| _objecttype_|Required;  **Variant** ( **String** ). The type or class of object to create.|
 **Remarks**
Use the  **GetObject** function to access an ActiveX object from a file and assign the object to an[object variable](vbe-glossary.md). Use the  **Set** statement to assign the object returned by **GetObject** to the object variable. For example:



```vb
Dim CADObject As Object
Set CADObject = GetObject("C:\CAD\SCHEMA.CAD")
```

When this code is executed, the application associated with the specified  **_pathname_** is started and the object in the specified file is activated.
If  **_pathname_** is a zero-length string (""), **GetObject** returns a new object instance of the specified type. If the **_pathname_** argument is omitted, **GetObject** returns a currently active object of the specified type. If no object of the specified type exists, an error occurs.
Some applications allow you to activate part of a file. Add an exclamation point ( **!** ) to the end of the file name and follow it with a string that identifies the part of the file you want to activate. For information on how to create this string, see the documentation for the application that created the object.
For example, in a drawing application you might have multiple layers to a drawing stored in a file. You could use the following code to activate a layer within a drawing called  `SCHEMA.CAD`:



```vb
Set LayerObject = GetObject("C:\CAD\SCHEMA.CAD!Layer3")
```

If you don't specify the object's  **_class_**, Automation determines the application to start and the object to activate, based on the file name you provide. Some files, however, may support more than one class of object. For example, a drawing might support three different types of objects: an **Application** object, a **Drawing** object, and a **Toolbar** object, all of which are part of the same file. To specify which object in a file you want to activate, use the optional **_class_** argument. For example:



```vb
Dim MyObject As Object
Set MyObject = GetObject("C:\DRAWINGS\SAMPLE.DRW", "FIGMENT.DRAWING")
```

In the example,  `FIGMENT` is the name of a drawing application and `DRAWING` is one of the object types it supports.
Once an object is activated, you reference it in code using the object variable you defined. In the preceding example, you access [properties](vbe-glossary.md) and[methods](vbe-glossary.md) of the new object using the object variable `MyObject`. For example:



```
MyObject.Line 9, 90
MyObject.InsertText 9, 100, "Hello, world."
MyObject.SaveAs "C:\DRAWINGS\SAMPLE.DRW"
```


 **Note**  Use the  **GetObject** function when there is a current instance of the object or if you want to create the object with a file already loaded. If there is no current instance, and you don't want the object started with a file loaded, use the **CreateObject** function.

If an object has registered itself as a single-instance object, only one instance of the object is created, no matter how many times  **CreateObject** is executed. With a single-instance object, **GetObject** always returns the same instance when called with the zero-length string ("") syntax, and it causes an error if the **_pathname_** argument is omitted. You can't use **GetObject** to obtain a reference to a class created with Visual Basic.

## Example

This example uses the  **GetObject** function to get a reference to a specific Microsoft Excel worksheet ( `MyXL`). It uses the worksheet's  **Application** property to make Microsoft Excel visible, to close it, and so on. Using two API calls, the `DetectExcel` **Sub** procedure looks for Microsoft Excel, and if it is running, enters it in the Running Object Table. The first call to **GetObject** causes an error if Microsoft Excel isn't already running. In the example, the error causes the `ExcelWasNotRunning` flag to be set to True. The second call to **GetObject** specifies a file to open. If Microsoft Excel isn't already running, the second call starts it and returns a reference to the worksheet represented by the specified file, mytest.xls. The file must exist in the specified location; otherwise, the Visual Basic error `Automation error` is generated. Next the example code makes both Microsoft Excel and the window containing the specified worksheet visible. Finally, if there was no previous version of Microsoft Excel running, the code uses the **Application** object's **Quit** method to close Microsoft Excel. If the application was already running, no attempt is made to close it. The reference itself is released by setting it to **Nothing**.


```vb
' Declare necessary API routines:
Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName as String, _
                    ByVal lpWindowName As Long) As Long

Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd as Long,ByVal wMsg as Long, _
                    ByVal wParam as Long, _
                    ByVal lParam As Long) As Long

Sub GetExcel()
    Dim MyXL As Object    ' Variable to hold reference
                                ' to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean    ' Flag for final release.

' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.
' Getobject function called without the first argument returns a 
' reference to an instance of the application. If the application isn't
' running, an error occurs.
    Set MyXL = Getobject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear    ' Clear Err object in case error occurred.

' Check for Microsoft Excel. If Microsoft Excel is running,
' enter it into the Running Object table.
    DetectExcel

' Set the object variable to reference the file you want to see.
    Set MyXL = Getobject("c:\vb4\MYTEST.XLS")

' Show Microsoft Excel through its Application property. Then
' show the actual window containing the file using the Windows
' collection of the MyXL object reference.
    MyXL.Application.Visible = True
    MyXL.Parent.Windows(1).Visible = True
     Do manipulations of your  file here.
    ' ...
' If this copy of Microsoft Excel was not running when you
' started, close it using the Application property's Quit method.
' Note that when you try to quit Microsoft Excel, the
' title bar blinks and a message is displayed asking if you
' want to save any loaded files.
    If ExcelWasNotRunning = True Then 
        MyXL.Application.Quit
    End IF

    Set MyXL = Nothing    ' Release reference to the
                                ' application and spreadsheet.
End Sub

Sub DetectExcel()
' Procedure dectects a running Excel and registers it.
    Const WM_USER = 1024
    Dim hWnd As Long
' If Excel is running this API call returns its handle.
    hWnd = FindWindow("XLMAIN", 0)
    If hWnd = 0 Then    ' 0 means Excel not running.
        Exit Sub
    Else                
    ' Excel is running so use the SendMessage API 
    ' function to enter it in the Running Object Table.
        SendMessage hWnd, WM_USER + 18, 0, 0
    End If
End Sub
```


