---
name: weekly-bookings-report
description: "Use when asked for the weekly bookings and cancellations report, or when it runs on its Monday schedule. Builds the report for the previous Monday to Sunday from the booking snapshot and posts the headline numbers."
---

# Weekly bookings and cancellations report

Ingrid Solberg's requirements are in the mail thread "Weekly bookings and cancellations report". This skill is written so that it can run unattended every Monday.

1. Call `backoffice_info` for today's date. The report week is the previous Monday to Sunday, by booking date.
2. Get the numbers from `bookings_summary` with `booked_from` and `booked_to` set to that week: once grouped by `seat_class`, once by `route`. Run it again for the week before, for the comparison.
3. Fill `templates/weekly-bookings-report-template.docx`: bookings, cancellations, revenue excluding cancelled bookings, the split by seat class, the three strongest routes by revenue.
4. "Anything unusual" is one line. A week-on-week change of more than 25% in bookings or cancellations counts as unusual. If nothing qualifies, write "nothing unusual".
5. Save a text version to the drive as `reports/weekly-bookings-<Monday date>.md`, and post the three headline numbers in #finance.
6. If any number could not be produced, say so in the post. Do not estimate.
