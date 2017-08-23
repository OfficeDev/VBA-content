---
title: "Свойство WebNavigationBarHyperlinks.Count (издатель)"
keywords: vbapb10.chm8585219
f1_keywords: vbapb10.chm8585219
ms.prod: publisher
api_name: Publisher.WebNavigationBarHyperlinks.Count
ms.assetid: 55e62f9b-7d7e-50bd-bd3b-0c2fdae903df
ms.date: 06/08/2017
ms.openlocfilehash: 8fd5b89e36bf2871277fcb1475fd0055b063706e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarhyperlinkscount-property-publisher"></a>Свойство WebNavigationBarHyperlinks.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **WebNavigationBarHyperlinks** .


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


