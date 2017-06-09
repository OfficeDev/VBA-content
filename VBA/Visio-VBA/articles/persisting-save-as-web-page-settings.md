---
title: Persisting Save as Web Page Settings
ms.prod: visio
ms.assetid: 3646a841-f99c-2906-856b-0fd5a642d499
ms.date: 06/08/2017
---


# Persisting Save as Web Page Settings

When a document is saved as a Web page with settings other than the default values described in the Save as Web Page API reference topics, selected settings are stored in the registry. These settings become the new default values until the properties are explicitly reset.

For example, if you do not want your files organized in a subfolder (the default) but prefer that all files be organized as flat files, set the  **StoreInFolder** property (or the **/folder** command-line option) to **False**. This setting becomes the default whenever you use the Save as Web Page feature.

This information is stored per user and is located in  **HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Visio\Solution\SaveAsWeb**.

The following Web page settings are persisted when their values are explicitly set to non-default values:

- altformat
    
- folder
    
- longnames
    
- navbar
    
- openbrowser
    
- panzoom
    
- priformat
    
- prop
    
- screenres
    
- search
    
- secformat
    
- stylesheet
    
- tabs
    
- theme
    
For information regarding default values, see the table describing command-line parameters in  [Running Save as Web Page from the command line](running-save-as-web-page-from-the-command-line.md) (the registry entries are the same as the command-line option names), or see the corresponding property topic. (The corresponding property topics are listed in [Running Save as Web Page from the command line](running-save-as-web-page-from-the-command-line.md).)

 **Note**  If for some reason the registry entries are corrupt or if you delete the  **SaveAsWeb** subkey in the registry, the solution reverts to using the original default values. These default values are stored internally in the solution and are used whenever the corresponding registry key does not exist.

Serious problems might occur if you modify the registry incorrectly by using Registry Editor or by using another method. These problems might require that you reinstall the operating system. Microsoft cannot guarantee that these problems can be solved. Modify the registry at your own risk.

