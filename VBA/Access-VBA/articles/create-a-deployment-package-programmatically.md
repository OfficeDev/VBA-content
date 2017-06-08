---
title: Create a Deployment Package Programmatically
ms.prod: access
ms.assetid: 4eb23608-e976-49a8-3f0e-f3537b948bfd
ms.date: 06/08/2017
---


# Create a Deployment Package Programmatically

The  **CreateInstallPackage** method enables you to create a deployment package programmatically.


## Syntax

 _expression_. **CreateInstallPackage**( ** _WizardSettingsFile_** )

 _expression_ A variable that represents an **AccessDeveloperExtensions** object.

The following table describes the arguments of the  **CreateInstallPackage** method.



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _WizardSettingsFile_|Required|**String**|The path and file name of the wizard template file that contains the settings to use when creating the deployment package. To create a wizard template (.adepsws), click  **Save Wizard Settings** on any Package Solution Wizard page.|

## Usage

You must instantiate the  **AccessDeveloperExtensions** object before you call the **CreateInstallPackage** method. Instantiating the **AccessDeveloperExtensions** object requires a different technique from instantiating the built-in objects in Access. To instantiate the **AccessDeveloperExtensions** object, you can use the **COMAddins** collection or the **CreateObject** method.

The following code illustrates how to instantiate the  **AccessDeveloperExtensions** object through the **COMAddins** collection.




```vb
Set objADE = Application.COMAddIns("AccessAddIn.ADE").Object 

```

The following code illustrates how to instantiate the  **AccessDeveloperExtensions** object by using the **CreateObject** method.




```vb
Set objADE = CreateObject("AccessAddIn.ADE") 

```

The following example wraps the steps necessary to call the  **CreateInstallPackage** method in a subroutine named **CreatePackage**. To use this example, pass the path and file name of the wizard template file to the subroutine. A deployment package will be created.




```vb
Sub CreatePackage(strSettingsPath As String) 
     
    Dim objAde As AccessDeveloperExtensions 
     
    ' Instantiate a AccessDeveloperExtensions object. 
    Set objAde = Application.COMAddIns("AccessAddIn.ADE").Object 
     
    ' Create the deployment package. 
    objAde.CreateInstallPackage strSettingsFilePath 
     
    Set objAde = Nothing 
End Sub
```

To use this example, you must set a reference to the Access Developer Extensions type library. To do this, follow these steps:


1. On the  **Tools** menu, click **References**.
    
2. Select the  **Microsoft Office Access Developer Extensions Type Library 1.0** check box, and then click **OK**.
    

