---
title: "Свойство Cell.MarginLeft (издатель)"
keywords: vbapb10.chm5111827
f1_keywords: vbapb10.chm5111827
ms.prod: publisher
api_name: Publisher.Cell.MarginLeft
ms.assetid: 1b665a3b-6958-0548-ece1-9d3a7045eaac
ms.date: 06/08/2017
ms.openlocfilehash: 7fa3868a5bc8fa45e81a3d17600ea249e02eb1fe
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellmarginleft-property-publisher"></a>Свойство Cell.MarginLeft (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текст и левой границей ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginLeft**

 переменная _expression_A, представляет собой объект- **ячейки** .


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


