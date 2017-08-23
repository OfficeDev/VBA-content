---
title: "Свойство TextFrame.MarginRight (издатель)"
keywords: vbapb10.chm3866646
f1_keywords: vbapb10.chm3866646
ms.prod: publisher
api_name: Publisher.TextFrame.MarginRight
ms.assetid: bdbde217-6a51-7823-ac93-8bbffa583544
ms.date: 06/08/2017
ms.openlocfilehash: a64bf65a16a9c3f5f38437ad9f1471da65606c31
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframemarginright-property-publisher"></a>Свойство TextFrame.MarginRight (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и правого края ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginRight**

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


