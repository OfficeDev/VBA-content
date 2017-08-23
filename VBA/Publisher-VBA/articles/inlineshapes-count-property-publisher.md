---
title: "Свойство InlineShapes.Count (издатель)"
keywords: vbapb10.chm5767171
f1_keywords: vbapb10.chm5767171
ms.prod: publisher
api_name: Publisher.InlineShapes.Count
ms.assetid: a78ae487-e7d6-1099-236f-6464c601686f
ms.date: 06/08/2017
ms.openlocfilehash: 6c870a4a3e4e11ef79bc3909ded2f6158fcf0063
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="inlineshapescount-property-publisher"></a>Свойство InlineShapes.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляющий объект **InlineShapes** .


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


