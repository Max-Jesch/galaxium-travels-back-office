---
owner: Jonas Keller
modified: 2026-10-15
---
# Runbook: booking API

## Payments time out
1. Check the payment provider's status page.
2. Check connection pool use on the dashboard. Above 80% for two minutes means we are queueing.
3. If the provider is slow, do not raise the pool further. Tell Support in #support-escalations that payments may fail and that retrying is safe.

## Bookings are rejected
- "Name does not match": working as designed. Send to Support.
- "Seat not available": check the hold service for stale holds on that flight. A hold older than 20 minutes is a bug.

## Who to wake
Primary on-call first. Marcus if customer money is involved. Farid if it lasts longer than an hour.
