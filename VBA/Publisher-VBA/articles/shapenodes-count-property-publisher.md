---
title: "Свойство ShapeNodes.Count (издатель)"
keywords: vbapb10.chm3473411
f1_keywords: vbapb10.chm3473411
ms.prod: publisher
api_name: Publisher.ShapeNodes.Count
ms.assetid: 5b259584-0aad-57bd-4848-cc7f6e96d430
ms.date: 06/08/2017
ms.openlocfilehash: 385945a47d34b35a5fe8fb0c2a0948e4b74e4b3d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodescount-property-publisher"></a>Свойство ShapeNodes.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **ShapeNodes** .


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


