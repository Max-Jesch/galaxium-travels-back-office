---
owner: Marcus Lindqvist
modified: 2099-10-02
---
# Booking platform: overview

```
 traveller ── web frontend ──┬── booking API (Python) ── database: users, flights, bookings
                             └── hold service (Java) ─── quotes and seat holds
                                        │
                                payment provider
```

| Part | What it does | Owner |
| --- | --- | --- |
| Web frontend | Search, quote, pay, manage a booking | Wren Abbott |
| Booking API | System of record for users, flights and bookings. Calculates the price of a booking when it is made. Applies the refund policy on cancellation. | Jonas Keller |
| Hold service | Since 1 October 2099. Issues a price quote and holds the seat for 15 minutes. | Aditi Sharma |
| Warehouse | Nightly copy of users, flights, bookings and quotes for reporting | Felix Braun |

## Things people get wrong
- A quote is not a booking. A quote becomes a booking only when the traveller pays inside the hold.
- The booking API rejects a booking when the traveller name does not match the account. This is deliberate.
- Seat classes are economy, business and galaxium. Each flight has 6, 3 and 1 seats.
- Charter flying is not in this system at all.

## Known follow-ups
See the roadmap in engineering/roadmap. The hold service's quote price is its own placeholder calculation and is to be replaced by the booking API's (issue 51).
