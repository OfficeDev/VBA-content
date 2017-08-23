---
title: "Свойство Window.Width (издатель)"
keywords: vbapb10.chm262150
f1_keywords: vbapb10.chm262150
ms.prod: publisher
api_name: Publisher.Window.Width
ms.assetid: 762df30a-7fdd-8f95-f64b-eae57e7c02fe
ms.date: 06/08/2017
ms.openlocfilehash: f5e9a582187433bcc8f348618a48e93d19fbcfcf
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowwidth-property-publisher"></a>Свойство Window.Width (издатель)

Возвращает или задает **Long** , представляющее ширину окна (в точках). Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Ширина**

 переменная _expression_A, представляющий объект **Window** .


## <a name="example"></a>Пример

Этот пример устанавливает высоту и ширину окна, если окно не развернуто и не свернуто.


```vb
Sub SetWindowHeight() 
 With ActiveWindow 
 If .WindowState = pbWindowStateNormal Then 
 .Height = InchesToPoints(5) 
 .Width = InchesToPoints(5) 
 End If 
 End With 
End Sub
```


