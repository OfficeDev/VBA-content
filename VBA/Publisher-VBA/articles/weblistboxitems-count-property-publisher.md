---
title: "Свойство WebListBoxItems.Count (издатель)"
keywords: vbapb10.chm4128771
f1_keywords: vbapb10.chm4128771
ms.prod: publisher
api_name: Publisher.WebListBoxItems.Count
ms.assetid: a306e5d1-c0e4-86f3-745a-720f91bf1f25
ms.date: 06/08/2017
ms.openlocfilehash: 6a247e0a5392c06f295633c1bb2d0d2dd5b1bdea
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weblistboxitemscount-property-publisher"></a>Свойство WebListBoxItems.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **WebListBoxItems** .


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


