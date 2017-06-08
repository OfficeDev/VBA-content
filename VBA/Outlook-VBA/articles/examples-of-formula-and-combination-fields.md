---
title: Examples of Formula and Combination Fields
ms.prod: outlook
ms.assetid: 40e7ff96-222b-14ce-102c-63467d2435f8
ms.date: 06/08/2017
---


# Examples of Formula and Combination Fields


|**Display**|**Custom field**|**Result in custom field**|
|:-----|:-----|:-----|
|Number of days since the item was received. ( **Formula** field)|DateValue (Now())-DateValue ([Received]) &; " Day(s)"|6 Day(s)|
|Description of a meeting or appointment in Calendar. ( **Formula** field)|"This meeting occurs " &; [Recurrence Pattern] &; " in " &; [Location]|The meeting occurs every day from 12:00 P.M. to 1:30 P.M. in room 1231|
|Amount to be charged for a phone call recorded in the Journal at $0.75 a minute. ( **Formula** field)|IIF ([Entry Type] = "Phone call" , Format ([Duration] * .75, "Currency"), "None")|$1.50|
|Description of a Message Flag. ( **Formula** field)|IIF ( [Flag Status] = "2", [Message Flag] &; " " &; [Due By],"")|Follow up 3/5/2009 8:00:00 A.M.|
|The first phone number recorded for a contact, in order of appearance in the formula. ( **Combination** field)|[Business Phone] [Business Phone 2] [Home Phone] [Home Phone 2] [Cell Phone]|(555) 555-1234|
|A description of a field combined with the field itself. ( **Combination** field)|Task Due: [Due Date]|Task Due: 3/5/2009 8:00:00 A.M.|

