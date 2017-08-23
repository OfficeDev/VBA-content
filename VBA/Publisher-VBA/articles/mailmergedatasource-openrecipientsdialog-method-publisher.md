---
title: "Метод MailMergeDataSource.OpenRecipientsDialog (издатель)"
keywords: vbapb10.chm6291490
f1_keywords: vbapb10.chm6291490
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.OpenRecipientsDialog
ms.assetid: 5a0a2b4a-ce23-435c-6e18-f778d6e14fd6
ms.date: 06/08/2017
ms.openlocfilehash: 37536b38c258d17e9f75950a7de2adf9f6b7c7b6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceopenrecipientsdialog-method-publisher"></a>Метод MailMergeDataSource.OpenRecipientsDialog (издатель)

Отображает диалоговое окно **Получатели** слияния почты публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OpenRecipientsDialog**

 переменная _expression_A, представляющий объект **вывода** .


## <a name="example"></a>Пример

В этом примере отображается диалоговое окно **Получатели слияния** .


```vb
Sub ShowRecipientsDialog() 
 ActiveDocument.MailMerge.DataSource.OpenRecipientsDialog 
End Sub
```


