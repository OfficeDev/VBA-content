---
title: Should I use a MultiPage or a TabStrip?
keywords: fm20.chm5225198
f1_keywords:
- fm20.chm5225198
ms.prod: office
ms.assetid: 3da861ef-58ca-1993-d661-b20c3d337673
ms.date: 06/08/2017
---


# Should I use a MultiPage or a TabStrip?

If you use a single layout for data, use a  **TabStrip** and map each set of data to its own **Tab**. If you need several layouts for data, use a **MultiPage** and assign each layout to its own **Page**.

Unlike a  **Page** of a **MultiPage**, the[client region](glossary-vba.md) of a **TabStrip** is not a separate form, but a portion of the form that contains the **TabStrip**. The border of a **TabStrip** defines a region of the form that you can associate with the tabs. When you place a control in the client region of a **TabStrip**, you are adding a control to the form that contains the **TabStrip**.

