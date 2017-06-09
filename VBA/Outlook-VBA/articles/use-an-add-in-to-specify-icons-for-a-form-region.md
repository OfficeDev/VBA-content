---
title: Use an Add-in to Specify Icons for a Form Region
ms.prod: outlook
ms.assetid: 7d542c9b-1881-780a-b58d-e34639399b60
ms.date: 06/08/2017
---


# Use an Add-in to Specify Icons for a Form Region

You can use an add-in to specify the custom icons you would like to use to help identify the state of an item in the explorer, inspector, and ribbon. Through the form region manifest XML file that you use to register the form region, you can specify the add-in that extends the form region, and the circumstances for which the custom icon is intended. When the specified cirsumstances occur, Outlook would obtain the appropriate icon from the add-in.


## To use an add-in to specify an icon for a form region


1. Implement the  **[FormRegionStartup](formregionstartup-object-outlook.md)** interface.
    
    All add-ins that extend form regions must implement the  **FormRegionStartup** interface. Outlook calls this interface to obtain layout storage data for a form region. For more information on add-ins for form regions, see [Extending a Form Region with an Add-in](extending-a-form-region-with-an-add-in.md).
    
    In particular, to specify custom icons, the add-in implements the  **[GetFormRegionManifest](formregionstartup-getformregionmanifest-method-outlook.md)** and the **[GetFormRegionIcon](formregionstartup-getformregionicon-method-outlook.md)** methods of the **FormRegionStartup** interface, specifying a form region manifest XML file and the circumstances where Outlook should display custom icons in the explorer, inspector, or ribbon. For example, you can create a form region to display a type of task that occurs in the household only, and these household tasks belong to a message class, **IPM.Task.Household**, which is derived from  **IPM.Task**. You can extend the form region with an add-in that specifies in the  **GetFormRegionIcon** method a special recurrent icon that Outlook should display adjacent to recurrent household tasks in the explorer.
    
2. In the form region manifest XML file, specify under the  **icons** element, the value `addin` for each of the child elements where you would like to use a custom icon.
    
    When Outlook displays items in the explorer or inspector, Outlook would look in the cache for the form region manifests that are associated with items of specific message classes. Where a child element of the  **icons** element has the value `addin`, Outlook calls  **GetFormRegionIcon** to obtain the corresponding icon and displays it accordingly for items of that message class.
    
    As an extension of the last example, in the form region manifest XML file for the form region associated with  **IPM.Task.Household**, you can specify under the  **icons** element, the value `addin` for the **recurring** child element. When Outlook displays all tasks in the explorer, Outlook would look at the cached form region manifest for items belonging to **IPM.Task.Household**. When Outlook realizes that the  **recurring** element has the value `addin`, Outlook will call  **GetFormRegionIcon** to obtain the icon for recurrent houshold tasks, and displays in the explorer the special recurrent icon adjacent to this type of task. For more information on child elements of the **icons** element, see [How to: Specify Icons to be Displayed for a Form Region](specify-icons-to-be-displayed-for-a-form-region.md).
    

