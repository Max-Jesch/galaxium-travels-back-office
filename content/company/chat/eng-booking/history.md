## 2099-09-24
[10:00] Marcus Lindqvist: Hold service goes live on 1 October. Freeze on the booking API from Monday, bug fixes only.
[10:20] Aditi Sharma: Load test passed at four times our busiest hour. Hold expiry job is the only thing I still want to watch in production.
[10:25] Clara Jensen: I start on the 5th, but I read the test plan. There is no test that compares the quote with the price of the booking. Is that on purpose?
    [10:31] Aditi Sharma: The quote price is a simple table for now. Wiring it to the booking backend's price calculation is the follow-up, issue 51.
    [10:33] Marcus Lindqvist: It is on the list for after launch.
## 2099-10-01
[08:40] Aditi Sharma: Deployed. Quotes and holds are on for everyone.
    + 🚀 9
[11:15] Wren Abbott: First quotes are coming in and converting. The countdown on the payment page works on every browser I can find.
## 2099-10-06
[14:00] Jonas Keller: Release 4.11 is out: fixes the timezone on booking confirmations, and the admin list of bookings now sorts by booking time.
## 2099-10-14
[09:30] Marcus Lindqvist: After yesterday: the payment timeout was a connection pool limit on the booking API, not the hold service. Jonas raised the limit and added an alert. Postmortem on Friday.
[09:45] Jonas Keller: Also fixed in 4.11.2: the confirmation mail job retried after a timeout and sent some customers two mails.
## 2099-10-21
[10:10] Maya Goldberg: Planning for release 4.12: refund preview before cancelling (issue 43), group holds (issue 47, design only), quote price from the booking backend (issue 51).
    [10:18] Marcus Lindqvist: 43 fits. 47 is first quarter. 51 depends on how much Aditi has left after the expiry job rework.
    [10:25] Aditi Sharma: 51 is two days of work and a week of testing. I would rather not rush it into 4.12.
    [10:27] Maya Goldberg: Then 4.13. I will put it there.
## 2099-10-29
[10:35] Aditi Sharma: Bug from Support: the hold timer starts on the server when the quote is created, so a slow page load eats into the 15 minutes. Proposal: start the timer when the quote is first shown. Issue 55.
    [10:50] Wren Abbott: I can send a "displayed" event from the page. Small change.
## 2099-11-03
[11:00] Marcus Lindqvist: Release 4.12 is planned for Tuesday 10 November. Code freeze on Friday.
[11:20] Clara Jensen: Test plan for 4.12 is in engineering/releases in the drive. I added the refund preview cases from the refund policy, including the tiers.
