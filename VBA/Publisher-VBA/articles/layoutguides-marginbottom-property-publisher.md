---
title: "Свойство LayoutGuides.MarginBottom (издатель)"
keywords: vbapb10.chm1114115
f1_keywords: vbapb10.chm1114115
ms.prod: publisher
api_name: Publisher.LayoutGuides.MarginBottom
ms.assetid: 9d11c4d9-8f53-7882-be40-200833a29fb6
ms.date: 06/08/2017
ms.openlocfilehash: 2fbc9cfdcf25882d662d12e042bc0f9fe17247f3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesmarginbottom-property-publisher"></a>Свойство LayoutGuides.MarginBottom (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и нижний край ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginBottom**

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


