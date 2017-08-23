---
title: "Свойство TextFrame.MarginLeft (издатель)"
keywords: vbapb10.chm3866644
f1_keywords: vbapb10.chm3866644
ms.prod: publisher
api_name: Publisher.TextFrame.MarginLeft
ms.assetid: 4e784b9f-9467-5a14-c211-589e69c3b8bc
ms.date: 06/08/2017
ms.openlocfilehash: 9552580e67f6c8992762b754453fae15fa07f51e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframemarginleft-property-publisher"></a>Свойство TextFrame.MarginLeft (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текст и левой границей ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginLeft**

 переменная _expression_A, представляет собой объект- **TextFrame** .


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


