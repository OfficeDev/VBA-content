---
title: "Свойство Cell.MarginRight (издатель)"
keywords: vbapb10.chm5111828
f1_keywords: vbapb10.chm5111828
ms.prod: publisher
api_name: Publisher.Cell.MarginRight
ms.assetid: d297222e-7fc1-9225-e098-1a85d7734d77
ms.date: 06/08/2017
ms.openlocfilehash: 67508f3bb58408977ff64dc8d8309d90de053c69
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellmarginright-property-publisher"></a>Свойство Cell.MarginRight (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и правого края ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginRight**

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


