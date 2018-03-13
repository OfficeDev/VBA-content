---
title: MousePointer Property
keywords: fm20.chm2001550
f1_keywords:
- fm20.chm2001550
ms.prod: office
api_name:
- Office.MousePointer
ms.assetid: ae574d87-e218-4d03-d423-0192768e82dc
ms.date: 06/08/2017
---


# MousePointer Property



Specifies the type of pointer displayed when the user positions the mouse over a particular object.
 **Syntax**
 _object_. **MousePointer** [= _fmMousePointer_ ]
The  **MousePointer** property syntax has these parts:


| <strong>Part</strong>   | <strong>Description</strong>                        |
|:------------------------|:----------------------------------------------------|
| <em>object</em>         | Required. A valid object.                           |
| <em>fmMousePointer</em> | Optional. The shape you want for the mouse pointer. |

 **Settings**
The settings for  _fmMousePointer_ are:


| <strong>Constant</strong>          | <strong>Value</strong> | <strong>Description</strong>                                                                                     |
|:-----------------------------------|:-----------------------|:-----------------------------------------------------------------------------------------------------------------|
| <em>fmMousePointerDefault</em>     | 0                      | Standard pointer. The image is determined by the object (default).                                               |
| <em>fmMousePointerArrow</em>       | 1                      | Arrow.                                                                                                           |
| <em>fmMousePointerCross</em>       | 2                      | Cross-hair pointer.                                                                                              |
| <em>fmMousePointerIBeam</em>       | 3                      | I-beam.                                                                                                          |
| <em>fmMousePointerSizeNESW</em>    | 6                      | Double arrow pointing northeast and southwest.                                                                   |
| <em>fmMousePointerSizeNS</em>      | 7                      | Double arrow pointing north and south.                                                                           |
| <em>fmMousePointerSizeNWSE</em>    | 8                      | Double arrow pointing northwest and southeast.                                                                   |
| <em>fmMousePointerSizeWE</em>      | 9                      | Double arrow pointing west and east.                                                                             |
| <em>fmMousePointerUpArrow</em>     | 10                     | Up arrow.                                                                                                        |
| <em>fmMousePointerHourglass</em>   | 11                     | Hourglass.                                                                                                       |
| <em>fmMousePointerNoDrop</em>      | 12                     | "Not" symbol (circle with a diagonal line) on top of the object being dragged. Indicates an invalid drop target. |
| <em>fmMousePointerAppStarting</em> | 13                     | Arrow with an hourglass.                                                                                         |
| <em>fmMousePointerHelp</em>        | 14                     | Arrow with a question mark.                                                                                      |
| <em>fmMousePointerSizeAll</em>     | 15                     | Size all cursor (arrows pointing north, south, east, and west).                                                  |
| <em>fmMousePointerCustom</em>      | 99                     | Uses the icon specified by the  <strong>MouseIcon</strong> property.                                             |

 **Remarks**
Use the  **MousePointer** property when you want to indicate changes in functionality as the mouse pointer passes over controls on a form. For example, the hourglass setting (11) is useful to indicate that the user must wait for a process or operation to finish.
Some icons vary depending on system settings, such as the icons associated with desktop themes.

