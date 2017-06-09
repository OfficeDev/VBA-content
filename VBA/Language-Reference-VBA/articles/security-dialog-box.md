---
title: Security Dialog Box
keywords: vbui6.chm181080
f1_keywords:
- vbui6.chm181080
ms.prod: office
ms.assetid: 2006719b-0e6f-47dc-4831-72a6ff205eb2
ms.date: 06/08/2017
---


# Security Dialog Box

Use this dialog box to determine the level of security used when opening documents, or to remove a certificate from the list of trusted sources.


## Security Level Tab

The settings on this tab indicate the level of security used when opening documents or loading add-ins.

 **High**

You can run code only in VBA projects that have been digitally signed and that are on your list of trusted sources (described below). If the certificate for a signed project is already on your list of trusted sources, it is automatically enabled and no warning is displayed. If the certificate for a signed project is not currently on your list of trusted sources, a warning is displayed and you can choose whether to enable or disable VBA code. If you choose to enable VBA code, you can choose to add the developer to the list of trusted sources. Before trusting a source, you should confirm that the source is responsible and uses a virus scanner before signing macros. Unsigned VBA projects are automatically disabled, and no warning is displayed. You cannot enable unsigned VBA projects at this security level.

 **Medium**

A warning is displayed whenever a VBA project from a source that is not on your list of trusted sources is loaded. You can choose whether to enable or disable both digitally signed and unsigned VBA projects. If the project might contain a virus, you should choose to disable the add-in. If the project has been signed, you can choose to add the developer to the list of trusted sources. Before trusting a source, you should confirm that the source is responsible and uses a virus scanner before signing macros.

 **Low**

If you are sure that all the VBA projects you load are safe, you can select this option — it turns off all virus protection. At this security level, VBA projects are always enabled.


## Trusted Sources Tab

This tab lists the currently trusted certificates that can be used by developers to sign documents and add-ins. When you open a digitally signed document, the digital signature appears on your computer as a certificate. The certificate names the VBA project's source, plus additional information about the identity and integrity of that source. A digital signature does not necessarily guarantee the safety of a project, and you must decide whether you trust a project that has been digitally signed. If you know you can always trust macros from a particular source, you can add that macro developer to the list of trusted sources when you open the project.


## Remove Button

If you added a certificate to your list of trusted sources when you first opened VBA project signed with that certificate, and later choose not to trust that source, you can use the  **Remove** button to remove the certificate from your list of trusted sources. The next time a project signed with that certificate is opened, the virus protection behavior corresponding to the setting on the **Security Level** tab will occur.


