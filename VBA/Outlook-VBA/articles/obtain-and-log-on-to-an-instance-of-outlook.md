---
title: Obtain and Log On to an Instance of Outlook
ms.prod: outlook
ms.assetid: ef369364-6500-2759-3ef4-ed4411112e96
ms.date: 06/08/2017
---


# Obtain and Log On to an Instance of Outlook

This topic shows how to obtain an  **[Application](application-object-outlook.md)** object that represents an active instance of Outlook, if there is one running on the local computer, or to create a new instance of Outlook, log on to the default profile, and return that instance of Outlook.

Helmut Obertanner provided the following code samples. Helmut is a [Microsoft Most Valuable Professional](https://mvp.microsoft.com/en-us/default.aspx) with expertise in Microsoft Office development tools in Microsoft Visual Studio and Microsoft Office Outlook.

The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. 

For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. 

You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code samples contain the  `GetApplicationObject` method of the `Sample` class, implemented as part of an Outlook add-in project. Each project adds a reference to the Outlook PIA, which is based on the **Microsoft.Office.Interop.Outlook** namespace.

The  `GetApplicationObject` method uses classes in the .NET Framework Class Library to check and obtain any Outlook process running on the local computer. It first uses the **[GetProcessesByName](http://msdn.microsoft.com/library/frlrfSystemDiagnosticsProcessClassGetProcessesByNameTopic%28Office.15%29.aspx)** method of the **Process** class in the **System.Diagnostics** namespace to obtain an array of process components on the local computer that share the process name "OUTLOOK". 

To check whether the array does contain at least one Outlook process, `GetApplicationObject` uses Microsoft Language Integrated Query (LINQ). The **[Enumerable](http://msdn.microsoft.com/library/frlrfSystemLinqEnumerableClassTopic%28Office.15%29.aspx)** class in the **[System.Linq](http://msdn.microsoft.com/library/frlrfSystemLinq%28Office.15%29.aspx)** namespace provides a set of methods, including the **[Count](http://msdn.microsoft.com/library/frlrfSystemLinqEnumerableClassCountTopic%28Office.15%29.aspx)** method, that implement the **[IEnumerable(T)](http://msdn.microsoft.com/library/frlrfSystemCollectionsGenericIEnumerable1ClassTopic%28Office.15%29.aspx)** generic interface. 

Because the **[Array](http://msdn.microsoft.com/library/frlrfSystemArrayClassTopic%28Office.15%29.aspx)** class implements the **IEnumerable(T)** interface, `GetApplicationObject` can apply the **Count** method to the array returned by **GetProcessesByName** to see whether there is an Outlook process running. If there is, `GetApplicationObject` uses the **[GetActiveObject](http://msdn.microsoft.com/library/frlrfSystemRuntimeInteropServicesMarshalClassGetActiveObjectTopic%28Office.15%29.aspx)** method of the **[Marshal](http://msdn.microsoft.com/library/frlrfSystemRuntimeInteropServicesMarshalClassTopic%28Office.15%29.aspx)** class in the **[System.Runtime.InteropServices](http://msdn.microsoft.com/library/frlrfSystemRuntimeInteropServices%28Office.15%29.aspx)** namespace to obtain that instance of Outlook, and casts that object to an Outlook **Application** object.

If Outlook is not running on the local computer, `GetApplicationObject` creates a new instance of Outlook, uses the **[Logon](namespace-logon-method-outlook.md)** method of the **[NameSpace](namespace-object-outlook.md)** object to log on to the default profile, and returns that new instance of Outlook.

The following is the C# code sample.


```C#
using System; 
using System.Diagnostics; 
using System.Linq; 
using System.Reflection; 
using System.Runtime.InteropServices; 
using Outlook = Microsoft.Office.Interop.Outlook; 
 
namespace OutlookAddIn1 
{ 
    class Sample 
    { 
        Outlook.Application GetApplicationObject() 
        { 
 
            Outlook.Application application = null; 
 
            // Check if there is an Outlook process running. 
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0) 
            { 
 
                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object. 
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application; 
            } 
            else 
            { 
 
                // If not, create a new instance of Outlook and log on to the default profile. 
                application = new Outlook.Application(); 
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI"); 
                nameSpace.Logon("", "", Missing.Value, Missing.Value); 
                nameSpace = null; 
            } 
 
            // Return the Outlook Application object. 
            return application; 
        } 
 
    } 
}
```

The following is the Visual Basic code sample.



```VB.net
Imports System.Diagnostics 
Imports System.Linq 
Imports System.Reflection 
Imports System.Runtime.InteropServices 
Imports Outlook = Microsoft.Office.Interop.Outlook 
 
Namespace OutlookAddIn2 
    Class Sample 
 
        Function GetApplicationObject() As Outlook.Application 
 
            Dim application As Outlook.Application 
 
            Check if there is an Outlook process running. 
            If Process.GetProcessesByName("OUTLOOK").Count() > 0 Then 
 
                ' If so, use the GetActiveObject method to obtain the process and cast it to an Application object. 
                application = DirectCast(Marshal.GetActiveObject("Outlook.Application"), Outlook.Application) 
            Else 
 
                ' If not, create a new instance of Outlook and log on to the default profile. 
                application = New Outlook.Application() 
                Dim ns As Outlook.NameSpace = application.GetNamespace("MAPI") 
                ns.Logon("", "", Missing.Value, Missing.Value) 
                ns = Nothing 
            End If 
 
            ' Return the Outlook Application object. 
            Return application 
        End Function 
 
    End Class 
End Namespace
```


