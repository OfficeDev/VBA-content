---
title: Security Behavior of the Outlook Object Model
ms.prod: outlook
ms.assetid: 4aa3b7c7-5f3f-41ce-bbf3-75d8ecbd6d4f
ms.date: 06/08/2017
---


# Security Behavior of the Outlook Object Model

## 

The Outlook object model includes entry points to access Outlook data, save data to specified locations, and send e-mails. These entry points are available to legitimate and malicious application developers alike. Versions of Outlook 98 and Outlook 2000 applied with the Outlook E-mail Security Update, and all subsequent versions starting from Outlook 2000 SP2 use the Object Model Guard to help protect users. 

The Object Model Guard warns users and prompts users for confirmation when untrusted applications attempt to use the object model to obtain e-mail address information, store data outside of Outlook, execute certain actions, and send e-mail messages. Although the Object Model Guard succeeds in identifying and protecting these entry points, two main issues exist that render the Object Model Guard rather unpractical:


- The default circumstances that applications invoke the Object Model Guard in earlier versions of Outlook can result in excessive security prompting for legitimate applications.
    
- The limitations of COM and Windows in identifying the specific application that is invoking the Object Model Guard have made it difficult for users to respond to the security prompts with certainty.
    
 For more information on the various security prompts of the Object Model Guard, see [Outlook Object Model Security Warnings](outlook-object-model-security-warnings.md). For more information on the protected object model entry points, see  [Protected Properties and Methods](protected-properties-and-methods.md).


## Default Security Behavior

Versions of Outlook prior to Outlook 2007 have relied on the Object Model Guard to protect Outlook address book data and avoid untrusted applications from sending e-mail. Although Outlook continues to use the Object Model Guard to provide similar protection, it has defined new default circumstances when the Object Model Guard generates warnings, reducing excessive security warnings under appropriate conditions while maintaining a reasonable degree of security for Outlook clients.

 **In-Process Add-ins**

In-process Outlook add-ins run in the process of the host Outlook program. In-process COM add-ins in Outlook are trusted by default. These COM add-ins are registered on the list of trusted applications by the client computer's administrator, and must use the  **[Application](application-object-outlook.md)** object that is passed to the **OnConnection** event of the add-in. Note that if you create a new **Application** object by using the **[CreateObject](application-createobject-method-outlook.md)** method, that object and any of its subordinate objects, properties, and methods are not trusted.

For more information about the  **OnConnection** event, see the [IDTExtensibility2](http://msdn.microsoft.com/library/frlrfExtensibilityIDTExtensibility2ClassTopic.aspx) documentation on MSDN.

 **Cross-Process Add-ins**

By default, Outlook relies on the existence and the status of an appropriate antivirus software on the client computer to trust cross-process applications: if Outlook detects that antivirus software is running with an acceptable status, Outlook will disable security warnings for the end user. All cross-process COM callers and add-ins will run without security warnings if all of the following conditions hold:


- The client computer is running Windows XP Service Pack 2 (SP2), Windows Vista, or a later version of Windows, and Windows Security Center (WSC) indicates that antivirus software on the computer is in a "Good" health status.
    
- The antivirus software installed on the client computer is designed for Windows XP SP2, Windows Vista, or later.
    
- Outlook is configured on the client computer in one of the following ways:
    
      - Uses the default Outlook security settings (that is, no Group Policy set up)
    
  - Uses security settings defined by Group Policy but does not have programmatic access policy applied
    
  - Uses security settings defined by Group Policy which is set to warn when the antivirus software is inactive or out of date
    
 For more information, see the "Code Security Changes in Microsoft Office Outlook 2007" article on MSDN.


## Security Options

 **Windows Group Policy**

Administrators can use the Trust Center in Outlook to change the default behavior. To access the Trust Center, select  **Tools** and then **Trust Center**. In the Trust Center, click  **Programmatic Access**. The  **Programmatic Access Security** dialog provides options other than the default behavior.

The three settings in the  **Programmatic Access Security** dialog are:


-  **Warn me about suspicious activity when my antivirus software is inactive or out-of-date (recommended)** This setting is the default, and implements the behavior described above. This is the recommended setting for all users.
    
-  **Always warn me about suspicious activity** This setting will revert Outlook to behave like Outlook 2003, where cross-process COM callers and untrusted add-ins will invoke security warnings.
    
-  **Never warn me about suspicious activity (not recommended)**This setting will never show security warnings and the Object Model Guard will be disabled. This setting should only be used in controlled environments where the risk of malicious code running on the computer is low.
    
These settings are only available if the current user is an administrator on the computer. Non-administrator users can see the current setting but will not be able to change it. Programmatic Access settings can also be controlled through Group Policy. For more information on configuring Outlook settings with Group Policy, see the Office Resource Kit Web site.

 **Security Form in Exchange Public Folder**

Administrators can configure Outlook to locate the Outlook security form in a public folder. In this case, Outlook will not leverage the status of antivirus software and will by default only trust add-ins listed in the security form. There will only be three prompt behaviors: prompt user, never prompt and automatically allow, and never prompt and automatically deny. 

To take advantage of the new code security behavior based on the status of antivirus software, administrators must use either the default Outlook security settings or configure Outlook to use Group Policy settings to override this behavior.


