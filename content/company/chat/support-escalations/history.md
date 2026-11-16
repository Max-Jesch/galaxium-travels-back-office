## 2026-10-14
[10:10] Thandi Dlamini: Escalation: a customer has two confirmation mails for one booking after yesterday's timeouts. One booking in the system. Can Engineering confirm there is no second charge?
    [10:40] Jonas Keller: Confirmed, one booking, one charge. The second mail was a retry of the mail job. Fix is in today's release.
## 2026-10-20
[13:15] Anneli Korhonen: Escalation to Operations: wheelchair assistance requests not reaching Tycho. Three cases, references in GT-1019.
    [15:00] Lucia Bianchi: Tycho says the requests arrive by mail to an address nobody reads at weekends. They have changed it.
## 2026-10-29
[09:50] Kofi Mensah: Escalation to Engineering: traveller sees "seat hold expired" immediately after requesting a quote. Twice for the same person, on a slow connection.
    [10:30] Aditi Sharma: The hold starts when the quote is created on the server, not when the page finishes loading. On a very slow connection that eats the 15 minutes. Noted as a bug.
