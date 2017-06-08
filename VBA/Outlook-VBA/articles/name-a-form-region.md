---
title: Name a Form Region
ms.prod: outlook
ms.assetid: 9e5009db-8230-3a82-60a6-d62cb5b0cc3c
ms.date: 06/08/2017
---


# Name a Form Region

Depending on the purpose, there are multiple ways you can provide identifiers for a form region. Only one of these identifiers is mandatory, which is the name that you associate with the form region manifest XML file when you register the form region in the Windows registry. You can optionally specify the other identifiers through the  **name**,  **title**, and  **formRegionName** elements of the form region XML schema.


## To provide an identifier in the Windows registry


1. In the Windows registry, under the current user key, create a key for the message class that the form region is created for, if the key does not already exist.
    
2. Add a value of the type  **String**.
    
3. Specify an identifier for the form region as the name of the key.
    
4. Specify the full path name of the form region manifest XML file as the data of the key.
    
For more information, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).


## Optional: To provide an internal identifier using the name element


- In the form region manifest XML file, specify a string identifier for the form region as a value of the  **name** element. The value of the **name** element is an internal name used by Outlook and any form region add-in to identify this form region. The following example specifies TestFormRegion as the internal identifier of a form region: `<name>TestFormRegion</name>`If the  **name** element is not specified in the form region manifest XML file, then the identifier specified in the Windows registry will be used as the internal name.
    

## Optional: To provide a display name using the title element


- In the form region manifest XML file, specify a string identifier for the form region as a value of the  **title** element. The value of the **title** element is the display name of the form region in the default locale. If the form region is a replacement or replace-all form region, the value of **title** is used in the **Actions** menu and the **Choose Form** dialog box. This value can be overridden with a locale-specific value according to the regional settings and any form region localization manifest. The following example specifies Sample Form Region as the display name of a form region: `<title>Sample Form Region</title>`In general, the value for the  **title** element is optional. If the **title** element is not specified in the form region manifest XML file, then the value for **formRegionName**, if it is defined, will be used as the display name. If no value for  **formRegionName** has been specified, then the internal name will be used. The exception is for a replacement or replace-all form region that has the **hidden** attribute being false. In this case, if the title element is not set, the form region will not be displayed in the **Actions** menu, **Choose Form** dialog, and **Design Form** dialog.
    

## Optional: To provide a form region identifier using the formRegionName element


- In the form region manifest XML file, specify a string identifier for the form region as a value of the  **formRegionName** element. The value of the **formRegionName** element identifies the form region in the **Show** group of the ribbon in the default locale. If the form region is an adjoining form region, the value is also used in the header that separates the beginning of an adjoining form region from the preceding portion of the form. This value can be overridden with a locale-specific value according to the regional settings and any form region localization manifest. The following example specifies Additional Information as the form region identifier of a separate form region: `<formRegionName>Additional Information</formRegionName>`If the  **formRegionName** element is not specified in the form region manifest XML file, then the value for **title**, if it is defined, will be used. If no value has been specified for  **title**, then the internal name will be used.
    

