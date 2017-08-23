---
title: "Свойство WebNavigationBarSets.Count (издатель)"
keywords: vbapb10.chm8454147
f1_keywords: vbapb10.chm8454147
ms.prod: publisher
api_name: Publisher.WebNavigationBarSets.Count
ms.assetid: ffe603c6-2c5a-de85-0924-aefa1dad269e
ms.date: 06/08/2017
ms.openlocfilehash: 34fa0112fe988fa387798e085a52e4e412800b0c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetscount-property-publisher"></a>Свойство WebNavigationBarSets.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSets** .


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


