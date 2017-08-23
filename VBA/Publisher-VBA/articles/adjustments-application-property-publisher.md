---
title: "Свойство Adjustments.Application (издатель)"
keywords: vbapb10.chm2424833
f1_keywords: vbapb10.chm2424833
ms.prod: publisher
api_name: Publisher.Adjustments.Application
ms.assetid: 9782bcd4-91ac-4ea3-4db7-f87b9b7c00ee
ms.date: 06/08/2017
ms.openlocfilehash: 739db05a0727091c7e9be172be8baa534002ecea
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="adjustmentsapplication-property-publisher"></a>Свойство Adjustments.Application (издатель)

При использовании без квалификатор объекта, данное свойство возвращает объект **[приложения](application-object-publisher.md)** , который представляет текущего экземпляра Publisher. Используется квалификатор объекта, данное свойство возвращает объект **приложения** , представляющего создателя указанный объект. При использовании с помощью объекта OLE-автоматизации возвращает объект приложения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Приложения**

 переменная _expression_A, представляющий объект **корректировки** .


## <a name="example"></a>Пример

В этом примере отображаются сведения о версии и построения для Publisher.


```vb
With Application 
 MsgBox "Current Publisher: version " _ 
 &; .Version &; " build " &; .Build 
End With
```

В этом примере отображается имя приложения, создавшего каждого связанного объекта на странице один активный публикации.




```vb
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект корректировки](adjustments-object-publisher.md)

