---
title: "Свойство CatalogMergeShapes.Count (издатель)"
keywords: vbapb10.chm8388611
f1_keywords: vbapb10.chm8388611
ms.prod: publisher
api_name: Publisher.CatalogMergeShapes.Count
ms.assetid: a871af2f-183c-f5a8-7ad0-c8d25c71e41f
ms.date: 06/08/2017
ms.openlocfilehash: 6870586f9802ea56dc5b1adca4466819b2d0d4b6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="catalogmergeshapescount-property-publisher"></a>Свойство CatalogMergeShapes.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **CatalogMergeShapes** .


## <a name="example"></a>Пример

В этом примере отображается число страниц в активный документ.


```vb
Sub CountNumberOfPages() 
 MsgBox "Your publication contains " &; _ 
 ActiveDocument.Pages.Count &; " page(s)." 
End Sub
```

В этом примере отображается количество фигур в активном документе.




```vb
Sub CountNumberOfShapes() 
 Dim intShapes As Integer 
 Dim pg As Page 
 
 For Each pg In ActiveDocument.Pages 
 intShapes = intShapes + pg.Shapes.Count 
 Next 
 
 MsgBox "Your publication contains " &; intShapes &; " shape(s)." 
End Sub
```


