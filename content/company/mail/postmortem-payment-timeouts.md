---
subject: "Postmortem: payment timeouts on 13 October"
labels: [engineering, incident]
attachments: [engineering/postmortems/2026-10-13-payment-timeouts.md]
---
From: Emeka Obi <emeka.obi@galaxium.example>
To: Engineering <engineering@galaxium.example>
Cc: Nadia Haddad <nadia.haddad@galaxium.example>, Farid Nazari <farid.nazari@galaxium.example>, Sofia Lindgren <sofia.lindgren@galaxium.example>
Date: 2026-10-16 12:15

All,

The postmortem for Tuesday's incident is in the drive. Summary: for 45 minutes the booking API could not confirm payments because every connection to the payment provider was waiting on a slow response. 37 payments timed out. No customer was charged twice. Some received two confirmation mails.

Actions and owners are in the document. Thank you to Support for the fast and accurate customer updates.

Emeka
