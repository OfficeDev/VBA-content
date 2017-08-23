---
title: "Свойство ColorFormat.RGB (издатель)"
keywords: vbapb10.chm2555904
f1_keywords: vbapb10.chm2555904
ms.prod: publisher
api_name: Publisher.ColorFormat.RGB
ms.assetid: aeff1962-b855-7c3f-1f4d-a336e0739ade
ms.date: 06/08/2017
ms.openlocfilehash: a4aae70b817f892a2905d5cbbfd537cdc90c4413
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorformatrgb-property-publisher"></a>Свойство ColorFormat.RGB (издатель)

Возвращает или задает **MsoRGBType** , который представляет значение красный зеленый синий (RGB) указанного цвета. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RGB**

 переменная _expression_A, представляет собой объект- **ColorFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoRGBType


## <a name="example"></a>Пример

В этом примере создается новая форма для первой страницы публикации, активных и задает красный цвет заливки.


```vb
Sub SetFill() 
 ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=100, Top:=100, Width:=100, Height:=100).Fill.ForeColor _ 
 .RGB = RGB(Red:=255, Green:=0, Blue:=0) 
End Sub
```

В этом примере возвращает значение цвет переднего плана первой фигуры на первой странице активных документов. В этом примере предполагает наличие по крайней мере один фигуры на первой странице active публикации.




```vb
Sub ShowFillColor() 
 MsgBox "The RGB fill value of this shape is " &; _ 
 ActiveDocument.Pages(1).Shapes(1).Fill.ForeColor.RGB &; "." 
End Sub
```


