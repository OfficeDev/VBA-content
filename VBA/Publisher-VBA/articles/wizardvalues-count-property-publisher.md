---
title: "Свойство WizardValues.Count (издатель)"
keywords: vbapb10.chm1638403
f1_keywords: vbapb10.chm1638403
ms.prod: publisher
api_name: Publisher.WizardValues.Count
ms.assetid: f32f3e88-fe3e-6d47-3579-c017e4fa2994
ms.date: 06/08/2017
ms.openlocfilehash: 6991e0772350385b3bbcc2696048e89437094ff4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardvaluescount-property-publisher"></a>Свойство WizardValues.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **WizardValues** .


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


