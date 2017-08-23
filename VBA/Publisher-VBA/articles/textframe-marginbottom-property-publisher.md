---
title: "Свойство TextFrame.MarginBottom (издатель)"
keywords: vbapb10.chm3866647
f1_keywords: vbapb10.chm3866647
ms.prod: publisher
api_name: Publisher.TextFrame.MarginBottom
ms.assetid: 55858bba-1103-48ba-64d6-5cc5ab677867
ms.date: 06/08/2017
ms.openlocfilehash: aff67c5c91ea4b92b16df6c5362fd183e25ee4b7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframemarginbottom-property-publisher"></a>Свойство TextFrame.MarginBottom (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и нижний край ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginBottom**

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


