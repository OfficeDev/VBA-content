---
title: CreateObject Function
keywords: vblr6.chm1010851
f1_keywords:
- vblr6.chm1010851
ms.prod: office
ms.assetid: d887c3d3-5c60-09a1-6856-49f7c4cc05ba
ms.date: 06/08/2017
---


# CreateObject Function



Creates and returns a reference to an [ActiveX object](vbe-glossary.md).
 **Syntax**
 **CreateObject(**_class,[servername]_**)**
The  **CreateObject** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _class_|Required;  **Variant** ( **String** ). The application name and class of the object to create.|
| _servername_|Optional;  **Variant** ( **String** ). The name of the network server where the object will be created. If _servername_ is an empty string (""), the local machine is used.|
The  _class_[argument](vbe-glossary.md) uses the syntax _appname_**.**_objecttype_ and has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _appname_|Required;  **Variant** ( **String** ). The name of the application providing the object.|
| _objecttype_|Required;  **Variant** ( **String** ). The type or[class](vbe-glossary.md) of object to create.|
 **Remarks**
Every application that supports Automation provides at least one type of object. For example, a word processing application may provide an  **Application** object, a **Document** object, and a **Toolbar** object.
To create an ActiveX object, assign the object returned by  **CreateObject** to an[object variable](vbe-glossary.md):



```vb
' Declare an object variable to hold the object 
' reference. Dim as Object causes late binding. 
Dim ExcelSheet As Object
Set ExcelSheet = CreateObject("Excel.Sheet")
```

This code starts the application creating the object, in this case, a Microsoft Excel spreadsheet. Once an object is created, you reference it in code using the object variable you defined. In the following example, you access [properties](vbe-glossary.md) and[methods](vbe-glossary.md) of the new object using the object variable, `ExcelSheet`, and other Microsoft Excel objects, including the  `Application` object and the `Cells` collection.



```vb
' Make Excel visible through the Application object.
ExcelSheet.Application.Visible = True
' Place some text in the first cell of the sheet.
ExcelSheet.Application.Cells(1, 1).Value = "This is column A, row 1"
' Save the sheet to C:\test.xls directory.
ExcelSheet.SaveAs "C:\TEST.XLS"
' Close Excel with the Quit method on the Application object.
ExcelSheet.Application.Quit
' Release the object variable.
Set ExcelSheet = Nothing

```

Declaring an object variable with the  `As Object` clause creates a variable that can contain a reference to any type of object. However, access to the object through that variable is late bound; that is, the binding occurs when your program is run. To create an object variable that results in early binding, that is, binding when the program is compiled, declare the object variable with a specific class ID. For example, you can declare and create the following Microsoft Excel references:



```vb
Dim xlApp As Excel.Application 
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.WorkSheet
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)

```

The reference through an early-bound variable can give better performance, but can only contain a reference to the [class](vbe-glossary.md) specified in the[declaration](vbe-glossary.md).
You can pass an object returned by the  **CreateObject** function to a function expecting an object as an argument. For example, the following code creates and passes a reference to a Excel.Application object:



```
Call MySub (CreateObject("Excel.Application"))
```

You can create an object on a remote networked computer by passing the name of the computer to the  _servername_ argument of **CreateObject**. That name is the same as the Machine Name portion of a share name: for a share named "\\MyServer\Public," _servername_ is "MyServer."

 **Note**  Refer to COM documentation (see  _Microsoft Developer Network_ ) for additional information on making an application visible on a remote networked computer. You may have to add a registry key for your application.

The following code returns the version number of an instance of Excel running on a remote computer named  `MyServer`:



```vb
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application", "MyServer")
Debug.Print xlApp.Version
```

If the remote server doesn't exist or is unavailable, a run-time error occurs.

 **Note**  Use  **CreateObject** when there is no current instance of the object. If an instance of the object is already running, a new instance is started, and an object of the specified type is created. To use the current instance, or to start the application and have it load a file, use the **GetObject** function.

If an object has registered itself as a single-instance object, only one instance of the object is created, no matter how many times  **CreateObject** is executed.

## Example

This example uses the  **CreateObject** function to set a reference ( `xlApp`) to Microsoft Excel. It uses the reference to access the  **Visible** property of Microsoft Excel, and then uses the Microsoft Excel **Quit** method to close it. Finally, the reference itself is released.


```vb
Dim xlApp As Object    ' Declare variable to hold the reference.
    
Set xlApp = CreateObject("excel.application")
    ' You may have to set Visible property to True
    ' if you want to see the application.
xlApp.Visible = True
    ' Use xlApp to access Microsoft Excel's 
    ' other objects.

```


