---
title: "Свойство ColorSchemes.Count (издатель)"
keywords: vbapb10.chm2752514
f1_keywords: vbapb10.chm2752514
ms.prod: publisher
api_name: Publisher.ColorSchemes.Count
ms.assetid: cd3fe69f-df35-8dcd-1133-634983155592
ms.date: 06/08/2017
ms.openlocfilehash: 58a52442c722711cb2b7b9ccd3127b4b9f7f9f11
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorschemescount-property-publisher"></a>Свойство ColorSchemes.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **ColorSchemes** .


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


