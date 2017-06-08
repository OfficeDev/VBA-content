---
title: Outlook Object Model Security Warnings
ms.prod: outlook
ms.assetid: 7e0cd805-5104-73af-d74f-b00480db91c4
ms.date: 06/08/2017
---


# Outlook Object Model Security Warnings

Subject to how Outlook has been configured to trust applications on a client computer, an application that uses the Outlook object model to access certain data or execute certain actions can invoke security warnings. Depending on the type of information or action that the program was attempting to access or execute, there are three different security prompts that applications can invoke through the Object Model Guard: the address book warning, send message warning, and execute action warning. This topic describes each of these security warnings. 

For more information on the default Outlook security behavior and security configuration options, see  [Security Behavior of the Outlook Object Model](security-behavior-of-the-outlook-object-model.md). For more information on the entry points in the object model that can trigger security warnings, see  [Protected Properties and Methods](protected-properties-and-methods.md).


## Address Book Warning

This warning is the most common security warning that is invoked when an untrusted application accesses Outlook data. Entry points that are identified with the "Address Book" prompt type in the topic  [Protected Properties and Methods](protected-properties-and-methods.md) can generate this warning.

This warning enables the user to allow or deny the action. The user can also choose to allow access to the address book for a period of time indicated in the drop-down box.

If the user clicks  **Deny**, Outlook immediately blocks the call that invokes the warning and returns  **MAPI_E_NOT_SUPPORTED**. Outlook does not return any data for the call. If the program does not properly handle the error, it might crash.

If the user clicks  **Allow** without selecting the **Allow access for** check box, only the call that generated the warning will be allowed. Additional calls on the same line or calls for objects that derive from the blocked call may generate their own security warnings.

If the user clicks  **Allow** after selecting the **Allow access for** check box, the call that generated the prompt, as well as future calls, will be allowed for the duration that the user has selected. During this time period, all callers to the object model, and not just the program that originally invoked the security warning, are approved for address-book access. After this time period expires, security warnings may reappear.


## Send Message Warning

This warning is invoked when an untrusted solution attempts to send an item programmatically. This dialog box has a built-in timer that prevents untrusted add-ins from sending messages rapidly and automatically. The user must wait five seconds before clicking  **Allow**.

If the user clicks  **Deny**, Outlook blocks the call that invoked the warning and returns the  **MAPI_E_NOT_SUPPORTED** error. Subsequent calls to send messages programmatically will invoke additional warnings.

If the user clicks  **Allow**, the call that invoked the warning, and only that call, is allowed. Subsequent calls from an untrusted solution to send messages programmatically will continue to generate warnings.


## Execute Action Warning

This warning is invoked when an untrusted solution executes a custom action from the  **[Actions](actions-object-outlook.md)** collection. Outlook displays a message similar to the previous warning, indicating that an action is being executed.

If the user clicks  **Deny**, Outlook blocks the call to the  **[Execute](action-execute-method-outlook.md)** method for that action and returns the **MAPI_E_NOT_SUPPORTED** error.

If the user clicks  **Allow**, the call that invoked the warning, and only that call, is allowed. Subsequent calls from an untrusted solution to execute an action will continue to invoke warnings.


