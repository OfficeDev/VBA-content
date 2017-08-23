---
title: "Свойство LayoutGuides.MarginRight (издатель)"
keywords: vbapb10.chm1114117
f1_keywords: vbapb10.chm1114117
ms.prod: publisher
api_name: Publisher.LayoutGuides.MarginRight
ms.assetid: 5dbfc999-59d6-c9d0-4d9d-bc1a4ee622aa
ms.date: 06/08/2017
ms.openlocfilehash: 6bf3603fcf20ef06bee92d117bedf1f712b1f8c6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesmarginright-property-publisher"></a>Свойство LayoutGuides.MarginRight (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и правого края ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginRight**

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


