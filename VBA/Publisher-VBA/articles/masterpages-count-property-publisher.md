---
title: "Свойство MasterPages.Count (издатель)"
keywords: vbapb10.chm589827
f1_keywords: vbapb10.chm589827
ms.prod: publisher
api_name: Publisher.MasterPages.Count
ms.assetid: adb14000-5dc4-9154-5c5f-8f63c89309b7
ms.date: 06/08/2017
ms.openlocfilehash: 55ce9773ad70222f6ed7ae60a7f3f061ee151a76
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="masterpagescount-property-publisher"></a>Свойство MasterPages.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **макетом** .


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


