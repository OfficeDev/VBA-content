---
title: Create Separate Read and Compose Pages Using Form Regions in a Custom Form
ms.prod: outlook
ms.assetid: 6e773aff-c7ec-f836-b4c2-84d6121fc62e
ms.date: 06/08/2017
---


# Create Separate Read and Compose Pages Using Form Regions in a Custom Form

To create distinct read and compose pages for a custom form, you can first use the Forms Designer to design a separate form region for the read page, and another separate form region for the compose page. For more information on creating a form region, see  [How to: Create a Form Region](create-a-form-region.md).

In order for Outlook to display the appropriate form region for the read page and the compose page, you can use an add-in that programmatically tells Outlook which form region to use in each case. Your add-in does this through the  **[GetFormRegionStorage](formregionstartup-getformregionstorage-method-outlook.md)** method of the **[FormRegionStartup](formregionstartup-object-outlook.md)** interface.

## To return the appropriate form region in the GetFormRegionStorage method


- In  **GetFormRegionStorage**, return the appropriate form region storage file (.OFS) based on the value that Outlook specifies for  _FormRegionMode_.
    
    As with any COM add-in that extends a form region, your add-in will implement the  **Outlook.FormRegionStartup** interface. In particular, when implementing the **GetFormRegionStorage** method of the **FormRegionStartup** interface, depending on the input value of the parameter _FormRegionMode_, your add-in will return the appropriate form region. For example, when Outlook calls  **GetFormRegionStorage** to obtain the form region for the read page, specifying _FormRegionMode_ as **olFormRegionRead**, your add-in will have implemented  **GetFormRegionStorage** to return the form .OFS file for the form region created for the read page. 

Similarly, when Outlook calls **GetFormRegionStorage** to obtain the form region for the compose page, specifying _FormRegionMode_ as **olFormRegionCompose**,  **GetFormRegionStorage** will return the local path to the .OFS file for the form region created for the compose page. Note that if your add-in specifies a path to an .OFS file as the return value for **GetFormRegionStorage**, the path must be a local path. For more information on implementing  **GetFormRegionStorage**, see  [Extending a Form Region with an Add-in](extending-a-form-region-with-an-add-in.md).
    

