---
title: "Свойство WizardProperty.Application (издатель)"
keywords: vbapb10.chm1572865
f1_keywords: vbapb10.chm1572865
ms.prod: publisher
api_name: Publisher.WizardProperty.Application
ms.assetid: 41d3b1aa-6d20-30e5-5736-cd94a9480951
ms.date: 06/08/2017
ms.openlocfilehash: cf56c45f7f47d7633a836b4079c4733f01d3d232
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardpropertyapplication-property-publisher"></a>Свойство WizardProperty.Application (издатель)

При использовании без квалификатор объекта, данное свойство возвращает объект **[приложения](application-object-publisher.md)** , который представляет текущего экземпляра Publisher. Используется квалификатор объекта, данное свойство возвращает объект **приложения** , представляющего создателя указанный объект. При использовании с помощью объекта OLE-автоматизации возвращает объект приложения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Приложения**

 переменная _expression_A, представляет собой объект- **WizardProperty** .


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


