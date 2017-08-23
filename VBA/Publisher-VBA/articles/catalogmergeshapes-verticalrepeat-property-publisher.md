---
title: "Свойство CatalogMergeShapes.VerticalRepeat (издатель)"
keywords: vbapb10.chm8388614
f1_keywords: vbapb10.chm8388614
ms.prod: publisher
api_name: Publisher.CatalogMergeShapes.VerticalRepeat
ms.assetid: 2a4852d6-14ee-7fa9-ea5e-170033c3a56d
ms.date: 06/08/2017
ms.openlocfilehash: 2c24c4e48c279b2d8c9260c7b5ee1a2d3c0cebd0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="catalogmergeshapesverticalrepeat-property-publisher"></a>Свойство CatalogMergeShapes.VerticalRepeat (издатель)

Возвращает значение типа **Long** , представляющее количество отправок области данных будет отображаться в конце конечной страницы публикации при выполнении объединение в каталог. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalRepeat**

 переменная _expression_A, представляет собой объект- **CatalogMergeShapes** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Когда выполняется объединение в каталог, область данных повторяется один раз для каждого выбранного записи в указанный источник данных.

Количество раз, когда область данных повторяет страницу вниз, определяется высота области. Свойство **[Height](shape-height-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** для возвращения или задания вертикальной размера области данных.

Свойство **[HorizontalRepeat](catalogmergeshapes-horizontalrepeat-property-publisher.md)** объекта **[CatalogMergeShapes](catalogmergeshapes-object-publisher.md)** представляет количество отправок повторении этой области по горизонтали между конечной страницы публикации.


## <a name="example"></a>Пример

В следующем примере возвращается количество раз, когда область данных будет повторяющиеся по горизонтали и по вертикали на конечной страницы публикации, когда выполняется объединение в каталог. В этом примере предполагается, что область данных является первой фигуры на первой странице указанной публикации.


```vb
Sub CatalogMergeDimensions() 
 
 With ThisDocument.Pages(1).Shapes(1) 
 Debug.Print .Width 
 Debug.Print .CatalogMergeItems.HorizontalRepeat 
 Debug.Print .Height 
 Debug.Print .CatalogMergeItems.VerticalRepeat 
 End With 
 
End Sub
```


