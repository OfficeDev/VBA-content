---
title: Customizing Outlook using COM add-ins
keywords: vbaol11.chm5272661
f1_keywords:
- vbaol11.chm5272661
ms.prod: outlook
ms.assetid: 84a4f616-3ace-0139-57d5-f0c070064ab2
ms.date: 06/08/2017
---


# Customizing Outlook using COM add-ins

Creating a COM add-in involves two major steps:


1. Implement the  **IDTExtensibility2** interface in a class module of a dynamic link library (DLL).
    
2. Register the COM add-in.
    

## Implement the IDTExtensibility2 Interface

The  **IDTExtensibility2** interface consists of five event procedures. To implement this interface in a Visual Basic program, set a reference to the Microsoft Add-In Designer object library and then add the following statement to the Declarations section of a class module:


```vb
Implements IDTExtensibility2
```

You can then add the empty event procedures to the code window of the class module and add your own program code to the procedures. You can also copy the empty procedures from an  [Outlook COM Add-in Template](outlook-com-add-in-template.md).


## Register the COM add-in

In order to work with Outlook, the add-in DLL must be registered. The DLL's class ID is registered beneath the \HKEY_CLASSES_ROOT subtree in the registry.

In addition, information about the add-in must be added to the registry. This information provides the add-in's name, description, target application, initial load behavior, and connection state.


 **Note**  If you use Microsoft Visual Basic 6.0 or later Developer to design your COM add-in, the add-in designer will perform the steps required to register the COM add-in for you.

The following example shows the contents of a sample registry-editor (.reg) file that illustrates how to register an Outlook COM add-in.




```text
[HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SampleAddIn.AddInIFace] 
"FriendlyName"="Sample Add-in" 
"Description"="Sample Outlook Add-In" 
"LoadBehavior"=dword:00000008
```

When the COM add-in is first registered,  **LoadBehavior** can be set to any of the following flags.



|**Value**|**Description**|
|:-----|:-----|
|2|Load at startup. The COM add-in is to be loaded and connected when Outlook starts.|
|8|Load on demand. The COM add-in is to be loaded and connected only when the user requests it, such as by using the  **COM Add-ins** dialog box.|
|16|Connect first time. The COM add-in is loaded and connected the first time the user runs Outlook after the COM add-in has been registered. The next time Outlook is run, the COM add-in is loaded when the user requests it. Use this value if your COM add-in modifies the user interface to allow the user to request the COM add-in be connected on demand (by clicking a button, for example).|
After the COM add-in is registered and loaded, the  **LoadBehavior** value can be combined with either of the following two flags to indicate current connection state of the COM add-in.



|**Flag**|**Description**|
|:-----|:-----|
|0|Disconnected|
|1|Connected|
To connect the COM add-in, set the Connected flag in  **LoadBehavior**; clear the flag to disconnect the COM add-in.

The  **FriendlyName** value specifies the name of the COM add-in as it's displayed in the **COM Add-in** dialog box. The **Description** value provides additional information about the COM add-in.


