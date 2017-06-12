---
title: VisDocCleanActions Enumeration (Visio)
keywords: vis_sdr.chm70310
f1_keywords:
- vis_sdr.chm70310
ms.prod: visio
ms.assetid: 78189c36-976b-6bcc-95fd-b38e2a74a285
ms.date: 06/08/2017
---


# VisDocCleanActions Enumeration (Visio)

Flags passed to the  **Document.Clean** method that indicate which document conditions to detect, report, and fix.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDocCleanActAll**|&;H3FFF|Perform all actions.|
| **visDocCleanActBadDisplayLists**|&;H100|Detect invalid display list linkages.|
| **visDocCleanActBadFieldFormulas**|&;H800|Detect fields that have missing or nonstandard formulas.|
| **visDocCleanActBadFieldMarks**|&;H1000|Detect fields that have out-of-sync count and marker values. Change the position of escape characters to match character counts.|
| **visDocCleanActBadReferences**|&;H2000|Detect formulas that have #Ref() errors.|
| **visDocCleanActConstantFormulas**|&;H20|Detect formulas that can be generated from the result.|
| **visDocCleanActDefault**|&;H1FD8|Default conditions to detect.|
| **visDocCleanActDeletedFields**|&;H400|Detect deleted fields.|
| **visDocCleanActDuplicateSubs**|&;H80|Detect duplicate subscriptions (cell dependencies).|
| **visDocCleanActEmptyRowsAndSects**|&;H2|Detect empty local rows and sections.|
| **visDocCleanActLocalFormulas**|&;H1|Detect unnecessary local overrides.|
| **visDocCleanActMissingSubs**|&;H10|Detect missing subscriptions (cell dependencies).|
| **visDocCleanActNearZero**|&;H40|Detect results that are almost zero and change them to zero.|
| **visDocCleanActNonDefaultFonts**|&;H4|Detect non-default font settings.|
| **visDocCleanActStaleResults**|&;H8|Detect results that don't match formulas.|
| **visDocCleanAlertDefault**|&;H0|Default conditions to report.|
| **visDocCleanFixDefault**|&;H3D8|Default conditions to fix.|

