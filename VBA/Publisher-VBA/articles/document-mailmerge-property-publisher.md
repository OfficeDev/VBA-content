---
title: "Свойство Document.MailMerge (издатель)"
keywords: vbapb10.chm196628
f1_keywords: vbapb10.chm196628
ms.prod: publisher
api_name: Publisher.Document.MailMerge
ms.assetid: 15b1a8aa-3472-c67d-1d99-92617b05c157
ms.date: 06/08/2017
ms.openlocfilehash: f662ee54b3788691b3d32a239624c0263187e53a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentmailmerge-property-publisher"></a>Свойство Document.MailMerge (издатель)

Возвращает объект **[слияния](mailmerge-object-publisher.md)** , который представляет функции слияния почты для указанной публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Слияния**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Слияния


## <a name="example"></a>Пример

В этом примере отображаются сведения из текущей записи в источнике данных.


```vb
Sub ViewMergeData() 
 ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False 
End Sub
```

В этом примере отображается диалоговое окно **Получатели слияния** , который содержит записи из источника данных.




```vb
Sub ExecuteMergeField() 
 ActiveDocument.MailMerge.DataSource.OpenRecipientsDialog 
End Sub
```


