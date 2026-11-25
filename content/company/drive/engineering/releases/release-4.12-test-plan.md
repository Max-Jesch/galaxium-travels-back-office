---
owner: Clara Jensen
modified: 2099-11-03
---
# Release 4.12: test plan

Release date: Tuesday 10 November 2099, 07:00 UTC. Code freeze Friday 6 November.

## Refund preview (issue 43)
The preview must show the same refund, fee and travel credit that the cancellation then applies. Cases are taken from policies/refund-policy.md, one per tier, plus:
- a booking that is already cancelled (no preview),
- a departure in the past (no preview),
- each seat class.

## Circuit breaker for the payment provider
- Provider slow: after five slow answers new payments fail fast with a clear message.
- Provider back: payments resume without a deploy.

## Hold timer starts on display (issue 55)
- Slow page load: the traveller still gets the full 15 minutes.
- The traveller never opens the page: the hold expires 20 minutes after the quote was created.

## Not in this release, not tested
Group holds (47). Quote price from the booking API (51).
