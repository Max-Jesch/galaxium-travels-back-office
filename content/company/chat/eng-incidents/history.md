## 2026-10-13
[15:52] Emeka Obi: **Incident opened.** Booking API response time above 8 seconds, payment confirmations timing out. I am incident lead.
[15:58] Jonas Keller: Database is fine. The API has no free connections to the payment provider, all 20 are waiting.
[16:10] Emeka Obi: Payment provider reports slow responses on their side since 15:45. We are queueing behind them.
[16:22] Jonas Keller: Raised the pool to 60 and shortened the timeout to 10 seconds. Deploying.
[16:35] Emeka Obi: Response time back under one second. Monitoring.
[17:05] Emeka Obi: **Incident closed.** Duration 45 minutes. 37 payments timed out, none charged twice. Postmortem Friday 16 October, 11:00.
## 2026-10-16
[12:10] Emeka Obi: Postmortem is in engineering/postmortems/2026-10-13-payment-timeouts.md. Actions: alert on pool use (done), circuit breaker for the payment provider (Jonas, 4.12), status page for Support (me, November).
    + 👍 7
## 2026-10-26
[09:00] Emeka Obi: On-call this week: Jonas primary, Aditi secondary. Rota for November is in engineering/on-call.
## 2026-11-02
[09:00] Emeka Obi: On-call this week: Aditi primary, Wren secondary. Nothing open from last week.
