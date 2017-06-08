---
title: Verify Whether Outlook Is a Click-to-Run Application on a Computer
ms.prod: outlook
ms.assetid: 4cdf9767-19b2-3976-460a-9470f5abac23
ms.date: 06/08/2017
---


# Verify Whether Outlook Is a Click-to-Run Application on a Computer

Click-to-Run is a software delivery and updating mechanism. Products delivered via Click-to-Run execute in a virtual application environment on the local operating system. This means that they have private copies of their files and settings, and that any changes they make are captured in the virtual environment. Click-to-Run is fast—users can start running an application within a short time without waiting for the entire product to finish installing. Updates are carried out automatically in the background, without requiring a user to first remove an installation or install patches. Click-to-Run products are virtualized and do not conflict with other installed software.

However, because a product delivered by Click-to-Run has private copies of all its files and registration, an add-in developer cannot determine the product's existence the same way as a product that has been installed on a client computer's hard drive. Starting in Office, Click-to-Run is the default mechanism to deliver Office, and only a subset of Office customers can request physical media to install Office. Add-in developers should determine whether Outlook has been installed, and whether Outlook has been delivered as a Click-to-Run product.

In Office, 32-bit Office and 64-bit Office are available via Click-to-Run. The default delivery is 32-bit Office for 32-bit or 64-bit Windows. You can also obtain 64-bit Office for a computer with 64-bit Windows. If you would like to have both Office 2010 and Office on the same computer, the bitness of the two versions of Office must be the same.

To check whether Outlook was delivered by Click-to-Run on a client computer:

- Verify if the  `VirtualOutlook` key exists in the following location in the Windows registry:
    
```text
HKEY_LOCAL_MACHINE\Software\Microsoft\Office\15.0\Common\InstallRoot\Virtual\VirtualOutlook
```

-  The `VirtualOutlook` key is a REG_SZ value that contains the culture tag of the installed product language, such as "en-us".
    


