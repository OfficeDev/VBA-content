---
title: WdTextboxTightWrap Enumeration (Word)
ms.prod: word
api_name:
- Word.WdTextboxTightWrap
ms.assetid: 6d9e0dcd-816a-e055-b96a-7da5dcea38f7
ms.date: 06/08/2017
---


# WdTextboxTightWrap Enumeration (Word)

Specifies how Microsoft Word tightly wraps text around text boxes.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdTightAll**|1|Wraps text around the text box tightly to the contents of the text box on all lines.|
| **wdTightFirstAndLastLines**|2|Wraps text tightly only on first and last lines.|
| **wdTightFirstLineOnly**|3|Wraps text tightly only on the first line.|
| **wdTightLastLineOnly**|4|Wraps text tightly only on the last line.|
| **wdTightNone**|0|Does not wrap text tightly around the contents of a text box.|

## Remarks

Typically, text is wrapped to the extents of the text box, including any white space around its contents. These settings allow the surrounding text to wrap to the contents of the text box and not the extents of the text box itself.


 **Note**  This setting works only if the text box is floating and has no border or fill set.


