---
title: Create a Database Template Programmatically
ms.prod: access
ms.assetid: fe4a1f39-a51b-b083-3673-095e5c6684e5
ms.date: 06/08/2017
---


# Create a Database Template Programmatically

The  **SaveAsTemplate** method enables you to convert an existing Access database file to a database template (.accdt) format file that can be featured on the **Getting Started with Microsoft Office Access** page.


## Syntax

 _expression_. **SaveAsTemplate**( **_TemplateLocation_**, **_TemplateName_**, **_PreviewImage_**, **_Description_**, **_Category_**, **_Keywords_**, **_Identifier_**, **_Reserved_** )

 _expression_ A variable that represents a **TemplateCreator** object.

The following table describes the arguments of the  **SaveAsTemplate** method.



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TemplateLocation_|Required|**String**|The full path and file name of the database template to create.For the template to appear on the  **Getting Started with Microsoft Office Access** page, it must be saved to **Microsoft\Templates** subfolder of the user's Application Data folder. In Windows XP, the default location of the Application Data folder is **C:\Documents and Settings\ _User Name_ \Application Data**, where _User Name_ is the name of the user who is currently logged on.In Windows Vista, the default location of the Application Data folder is **C:\Users\ _User Name_ \AppData\Roaming**, where _User Name_ is the name of the user who is currently logged on.You can use the **Environ** function to determine the current location of the user's Application Data folder. The following code illustrates how to do this. `strTemplateLocation = Environ("AppData") &; "\Microsoft\Templates\"`|
| _TemplateName_|Optional|**String**|The name of the database that is created when the user opens the template.|
| _PreviewImage_|Optional|**String**|An image file to be used as a preview for the database template on the  **Getting Started with Microsoft Office Access** page.|
| _Description_|Optional|**String**| A description to be displayed when the user selects the database template in the **Getting Started with Microsoft Office Access** page.|
| _Category_|Optional|**String**|The  **Template Category** under which the database template will appear on the **Getting Started with Microsoft Office Access** page.|
| _Keywords_|Optional|**String**|Keywords to be added to the template's file properties.|
| _Identifier_|Optional|**String**||
| _Reserved_|Optional|**String**||

## Usage

You must instantiate the  **TemplateCreator** object before you call the **SaveAsTemplate** method. Instantiating the **TemplateCreator** object requires a different technique from instantiating the built-in objects in Access. To instantiate the **TemplateCreator** object, you must use the **COMAddins** collection.

The following code illustrates how to instantiate the  **AccessDeveloperExtensions** object through the **COMAddins** collection.




```vb
Set objTemplate = Application.COMAddIns("AccessAddIn.ADE").Object.TemplateObject 

```

The following example creates a new template named Asset Tracker and assigns it to the Departmental data category on the  **Getting Started with Microsoft Office Access** page.




```vb
    Dim objTemplate As TemplateCreator 
    Dim strTemplateLocation As String 
     
    ' The database template must be saved to this location to appear on the 
    ' Getting Started with Microsoft Office Access page. 
    strTemplateLocation = Environ("AppData") &; "\Microsoft\Templates\" 
     
    ' Instantiate a TemplateObject object. 
    Set objTemplate = Application.COMAddIns("AccessAddIn.ADE").Object.TemplateObject 
 
    ' Create the database template.     
    objTemplate.SaveAsTemplate TemplateLocation:=strTemplateLocation &; "AssetTracker.accdt", _ 
                               TemplateName:="Asset Tracker", _ 
                               Category:="Departmental Data"
```

You must set a reference to the Access Developer Extensions type library in order to use the  **SaveAsTemplate** method. To do this, follow these steps:


1. On the  **Tools** menu, click **References**.
    
2. Select the  **Microsoft Office Access Developer Extensions Type Library 1.0** check box, and then click **OK**.
    



