---
title: "Свойство CellRange.Count (издатель)"
keywords: vbapb10.chm5177347
f1_keywords: vbapb10.chm5177347
ms.prod: publisher
api_name: Publisher.CellRange.Count
ms.assetid: b21dfbc8-fa1d-aa25-c8a2-ed81629b5da1
ms.date: 06/08/2017
ms.openlocfilehash: 5f16b25d331e8c2f30a4aa76d447957eee2fab8f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellrangecount-property-publisher"></a>Свойство CellRange.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **CellRange** .


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


