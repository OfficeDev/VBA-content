---
title: Constants and Variables in VBScript
keywords: olfm10.chm3077515
f1_keywords:
- olfm10.chm3077515
ms.prod: outlook
ms.assetid: f04a4521-5bb9-39e0-f7e2-a2b74193b827
ms.date: 06/08/2017
---


# Constants and Variables in VBScript

In VBScript, constants must be referenced by their numeric values. The constant string does not work and returns a value of 0, which gives unpredictable results.

There are two types of variables. Procedure-level variables that are used only within a procedure and script-level variables that are available to all the procedures within your script. Declare script-level variables at the top of your script. Declare procedure-level variables inside procedures. You can use procedure-level variables with the same name in different procedures because each variable is recognized only by the procedure in which it's declared. When the procedure exits, the variable ends. Variables that refer to Outlook objects can be either procedure-level or script-level variables. However, the value of the variable must be set within a procedure. Do not attempt to access Outlook objects outside of a procedure.

## Rules about variables:


- Must begin with an alphanumeric character.
    
- Cannot contain an embedded period.
    
- Cannot exceed 255 characters.
    
- Cannot use more than 127 procedure-level variables (arrays count as a single variable).
    
- Cannot use more than 127 script-level variables.
    

