---
title: "Свойство TextFrame.MarginTop (издатель)"
keywords: vbapb10.chm3866645
f1_keywords: vbapb10.chm3866645
ms.prod: publisher
api_name: Publisher.TextFrame.MarginTop
ms.assetid: 9709eefe-0857-f228-aa56-780c4789a413
ms.date: 06/08/2017
ms.openlocfilehash: 46f89ae4779febd39144d207bf393914590ce228
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframemargintop-property-publisher"></a>Свойство TextFrame.MarginTop (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и верхнего края ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginTop**

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


