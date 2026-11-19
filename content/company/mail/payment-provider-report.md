---
subject: Incident report 13 October, degraded authorisation times
labels: [engineering, vendor, incident]
---
From: Status <status@ferrypay.example>
To: Emeka Obi <emeka.obi@galaxium.example>
Date: 2026-10-15 09:00

Dear customer,

On 13 October between 15:45 and 16:31 UTC, authorisation requests took up to 25 seconds because of a failed database failover in our primary region. No authorisations were lost or duplicated. We have corrected the failover configuration and will run a test on 27 October at 03:00 UTC, which you should not notice.

Ferrypay Operations
