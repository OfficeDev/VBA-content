---
title: "Свойство WizardProperties.Count (издатель)"
keywords: vbapb10.chm1507331
f1_keywords: vbapb10.chm1507331
ms.prod: publisher
api_name: Publisher.WizardProperties.Count
ms.assetid: 835f3467-ec89-54d2-c685-3021e6267121
ms.date: 06/08/2017
ms.openlocfilehash: 777825e3a52212716b177c1c5385f2b661c3e70b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardpropertiescount-property-publisher"></a>Свойство WizardProperties.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **WizardProperties** .


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


