---
title: Best Practices to Migrate Outlook 97-2003 Custom Forms
ms.prod: outlook
ms.assetid: dd1170f1-ff20-12ab-7bf8-81df434ef143
ms.date: 06/08/2017
---


# Best Practices to Migrate Outlook 97-2003 Custom Forms

In general, a custom form solution that was built with Microsoft Office Outlook 2003 or earlier and that has been published to a forms library or Outlook folder are still supported. If you are used to sending a custom form definition with items (as one-off forms) using Outlook 2003 or an earlier version of Outlook, there will be circumstances where this will fail, for example, if you use Microsoft Exchange Server 2007. In such cases, instead of sending the custom form definition, you should make the form available by publishing it to a forms library or an Outlook folder.

If you want to further customize your Outlook 97-2003 custom form to take advantage of form regions, it is best that you redesign the entire solution to use only form regions. In particular, if your Outlook 97-2003 custom form is associated with a custom message class, you should deprecate that form and associate the custom message class with the new form region solution.

However, if for some reason, you cannot implement some existing functionality using form regions, you should keep the existing custom form pages together with any of their custom actions as intact as possible, and implement only new functionality using form regions. In this case, you should add new actions as custom actions only to the form regions.


