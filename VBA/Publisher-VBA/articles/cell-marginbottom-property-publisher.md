---
title: "Свойство Cell.MarginBottom (издатель)"
keywords: vbapb10.chm5111826
f1_keywords: vbapb10.chm5111826
ms.prod: publisher
api_name: Publisher.Cell.MarginBottom
ms.assetid: a05fd3a4-f4d5-232a-1f5d-0fa1bce136bd
ms.date: 06/08/2017
ms.openlocfilehash: 42a36d54af1f0c96b56a8ebe73c27dfe3a0e6452
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellmarginbottom-property-publisher"></a>Свойство Cell.MarginBottom (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и нижний край ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginBottom**

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


