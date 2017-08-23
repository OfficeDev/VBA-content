---
title: "Свойство LayoutGuides.MarginLeft (издатель)"
keywords: vbapb10.chm1114116
f1_keywords: vbapb10.chm1114116
ms.prod: publisher
api_name: Publisher.LayoutGuides.MarginLeft
ms.assetid: 02d1a544-3e41-3875-3027-61bdc465e89b
ms.date: 06/08/2017
ms.openlocfilehash: 34d48d4ca647ad570d8ed31113c5769bc4546a4f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesmarginleft-property-publisher"></a>Свойство LayoutGuides.MarginLeft (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текст и левой границей ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginLeft**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


## <a name="example"></a>Пример

В этом примере задается полей активной публикации для двух дюйма.


```vb
Sub SetPageMargins() 
 
 With ActiveDocument.LayoutGuides 
 .MarginTop = Application.InchesToPoints(Value:=2) 
 .MarginBottom = Application.InchesToPoints(Value:=2) 
 .MarginLeft = Application.InchesToPoints(Value:=2) 
 .MarginRight = Application.InchesToPoints(Value:=2) 
 End With 
 
End Sub
```


