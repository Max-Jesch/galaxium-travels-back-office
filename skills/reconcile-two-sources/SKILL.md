---
name: reconcile-two-sources
description: "Use when asked to reconcile, compare or check two sets of records against each other, for example quotes against bookings or a spreadsheet tracker against the CRM. Produces counts, totals, the cause and a list of what could not be verified."
---

# Reconcile two sources

This is how Galaxium's finance team reconciles. Follow the order.

1. **Name the system of record.** Decide which source is authoritative before comparing (the booking system for bookings, the CRM for deals). Say so in the result.
2. **Check each source against itself first.** Count rows, look for duplicate ids and blank keys. An export pulled by hand is not trusted until its row count and total agree with the system it came from.
3. **Join on the id, never on names or amounts.** State the key you used and how many records had no partner on either side.
4. **Report in this order:** how many match, how many differ in each direction, the total of each direction, the net. Then the split the requester asked for.
5. **Look for the rule behind the differences.** Compare against the written policy in the drive. A cause is only stated as fact when every differing record fits it; otherwise call it a hypothesis and give the count that fits.
6. **Finish with two lists:** the source of every figure, and the things you assumed or could not verify.

Write the result with `templates/memo-template.md`. Do not round amounts. Do not correct records yourself; a system fault goes to the engineering queue as a ticket.
