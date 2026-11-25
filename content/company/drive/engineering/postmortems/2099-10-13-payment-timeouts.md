---
owner: Emeka Obi
modified: 2099-10-16
---
# Postmortem: payment timeouts, 13 October 2099

**Duration.** 15:50 to 16:35 UTC, 45 minutes.
**Impact.** 37 payments timed out. No customer was charged twice. 23 customers received two confirmation mails. Support handled 47 contacts.

## Timeline (UTC)
- 15:45 The payment provider's authorisations slow down (their failed database failover).
- 15:50 Our booking API runs out of connections to the provider; all 20 are waiting.
- 15:52 Alert on response time. Incident opened.
- 16:22 Pool raised to 60, timeout shortened to 10 seconds. Deployed.
- 16:35 Response times normal.
- 17:05 Incident closed.

## Causes
1. The provider was slow.
2. Our pool was small and our timeout was long, so a slow provider stopped every booking.
3. The confirmation mail job retried after a timeout without checking whether the first mail had gone.

## Actions
| Action | Owner | Due | Status |
| --- | --- | --- | --- |
| Alert on connection pool use | Jonas | 14 Oct | Done |
| Fix the mail retry | Jonas | 14 Oct | Done (4.11.2) |
| Circuit breaker for the payment provider | Jonas | Release 4.12 | In progress |
| Status page that Support can point customers to | Emeka | November | Open |

No blame. The pool size had been the same since 2097 and nobody had a reason to look at it.
