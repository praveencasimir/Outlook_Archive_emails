# Outlook Archive mails

This VBA will help in below scenario.
Your org enforces your default inbox retention policy to a mininum of 30 or 90 days and if you have to move mails manually to your 3 year or longer retention period folders.

### How To
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
│   │   ├── Alerts-test
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

## Steps performed (every 1 hour)
* MoveConversationHistoryItems (all)
* MoveAlertItems (all)
* MoveAlertTestItems (all)
* MoveDeletedItems (> 30 days)
* MoveSentItems (> 30 days)
* MoveInboxItems (> 30 days)
