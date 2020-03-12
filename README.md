# OutlookRules

* Open `Outlook`
* Open `Files` Menu ->`Options`->`Customize Ribbon`
* On the right side enable `Developer` check box -> press `OK`
* Open `Developers` Menu in the ribbon and click on `Visual Basic`
* A default project would open if you have not entered any scripts earlier
* Expand Project1 -> Microsoft Outlook Objects -> This Outlook Session
* Copy contents of [ThisOutlookSession](ThisOutlookSession) and paste it in This Outlook Session code area.
* Create new Module and place contents of [OutlookMoveFilesVBA](OutlookMoveFilesVBA) in it.

The code caters to below mailbox structure.

```
mailbox/
├── Inbox
│   ├── Alerts
│   │   ├── Alerts-Com2passtest
│   ├── Junk Email
│   ├── Scanned
├── Conversation History
├── Deleted Items
├── __3 Years(36 months)
│   ├── Inbox
│   ├── Alerts
│   ├── Drafts
│   ├── Sent Items
│   ├── ConversationHistory
```

## Steps performed
* MoveConversationHistoryItems
* MoveAlertItems
* MoveAlertCompassItems
* MoveDeletedItems
* MoveSentItems
* MoveInboxItems
