---
title: "Свойство MailMerge.WizardState (издатель)"
keywords: vbapb10.chm6225929
f1_keywords: vbapb10.chm6225929
ms.prod: publisher
api_name: Publisher.MailMerge.WizardState
ms.assetid: a237cb3f-2c03-5f62-fa67-d4aa7703389d
ms.date: 06/08/2017
ms.openlocfilehash: c948ff59239ce320fe11f25891636c0dc7a699a2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergewizardstate-property-publisher"></a>Свойство MailMerge.WizardState (издатель)

Возвращает или задает **Long** , указывающее, текущий этап мастера слияния для публикации. Свойство **WizardState** возвращает номер, который указывает на текущий этап мастер слияния почты; нуль (0) означает, что закрытия мастера слияния почты. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WizardState**

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере отображается мастер слияния почты, если он работает.


```vb
Sub ShowMergeWizard() 
 With ActiveDocument.MailMerge 
 If .WizardState = 0 Then 
 .ShowWizard 
 End If 
 End With 
End Sub
```


