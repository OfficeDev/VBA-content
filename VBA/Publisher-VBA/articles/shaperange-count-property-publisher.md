---
title: "Свойство ShapeRange.Count (издатель)"
keywords: vbapb10.chm2293763
f1_keywords: vbapb10.chm2293763
ms.prod: publisher
api_name: Publisher.ShapeRange.Count
ms.assetid: 5037bfe9-b430-4205-c514-b2f4313b4c53
ms.date: 06/08/2017
ms.openlocfilehash: 992e8b0d09c458f27369ec7e14874c0ec23968ad
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangecount-property-publisher"></a>Свойство ShapeRange.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


