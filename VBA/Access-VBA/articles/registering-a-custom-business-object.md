---
title: Registering a Custom Business Object
ms.prod: access
ms.assetid: eed3b78e-310a-98fa-5cf9-32edaab0402f
ms.date: 06/08/2017
---


# Registering a Custom Business Object

  

**Applies to:** Access 2013 | Access 2016

To successfully launch a custom business object (.dll or .exe) through the Web server, the business object's ProgID must be entered into the registry as explained in this procedure. This RDS feature protects the security of your Web server by running only sanctioned executables.


 **Note**  For MDAC 2.0 and later, the default business object,  **RDSServer.DataFactory**, is not registered by default during MDAC installation. However, if **RDSServer.DataFactory** was registered as safe for execution on the computer prior to the installation, the registry entry is maintained for the new installation.

 **To register a custom business object**

1. Click  **Start** and then click **Run**.
    
2. Type  **RegEdit** and click **OK**.
    
3. In the Registry Editor, navigate to the  **HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\W3SVC\Parameters\ADCLaunch** registry key.
    
4. Select the  **ADCLaunch** key, and then from the **Edit** menu, point to **New** and click **Key**.
    
5. Type the ProgID of your custom business object and click  **Enter**. Leave the **Value** entry blank.
    
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

