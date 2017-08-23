---
title: "Свойство LayoutGuides.MarginTop (издатель)"
keywords: vbapb10.chm1114118
f1_keywords: vbapb10.chm1114118
ms.prod: publisher
api_name: Publisher.LayoutGuides.MarginTop
ms.assetid: f0b4f600-6c79-060b-edd5-82f07f78770a
ms.date: 06/08/2017
ms.openlocfilehash: eecbf5e1d30bc77c7f648ec6db166e7b596d182d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesmargintop-property-publisher"></a>Свойство LayoutGuides.MarginTop (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и верхнего края ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginTop**

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


