# Plan: Quote-to-Booking Join & Mismatch Read Tool

## Top-Level Overview

Add a read-only MCP tool (`bookings_quote_mismatches`) to the Galaxium Back Office MCP server under the `bookings` system. This tool joins converted price quotes from the hold service to their resulting bookings, compares `quoted_price` against `price_paid`, identifies pricing discrepancies (overcharged / undercharged), and provides summary counts and totals alongside paginated, filterable joined records.

---

## Sub-Tasks

### Sub-Task 1: Core Join & Mismatch Calculation Logic
- **Intent**: Implement the data join between `quotes` and `bookings` snapshots, calculate price differences, enrich records with route and flight details, and compute aggregate mismatch metrics.
- **Expected Outcomes**:
  - Converted quotes (`status == "converted"`) are matched with their corresponding bookings via `booking_id`.
  - Price difference (`price_paid - quoted_price`) is calculated for each pair.
  - Summary metrics are computed: total joined pairs, mismatch count, count and total amount charged more (overcharged), count and total amount charged less (undercharged), and matching count.
- **Todo List**:
  1. Retrieve `bookings`, `quotes`, and `flights` tables from `store.snapshot()`.
  2. Index bookings and flights for efficient join lookup.
  3. Iterate through converted quotes and join with bookings by `booking_id`.
  4. Enrich records with route info (`origin to destination`), quoted price, price paid, and signed difference.
  5. Compute aggregate summary metrics for overcharged, undercharged, and exact matches.
- **Relevant Context**:
  - `backoffice/store.py`: `snapshot()` table loader.
  - `backoffice/server.py`: `_route_lookup()`, `_page()`, `@_read("bookings")`.
- **Status**: `[ ] pending`

---

### Sub-Task 2: Tool Interface, Filtering & Pagination
- **Intent**: Expose the tool via the MCP server interface with support for query filters and pagination.
- **Expected Outcomes**:
  - The `@_read("bookings")` decorator registers `bookings_quote_mismatches` with read-only annotations.
  - Users and agents can filter by `user_id`, `flight_id`, `seat_class`, and toggle `mismatch_only` (default: True).
  - Paged items are returned using `_page(items, limit, offset)` alongside the overall summary metrics.
- **Todo List**:
  1. Define function signature for `bookings_quote_mismatches(user_id=0, flight_id=0, seat_class="", mismatch_only=True, limit=50, offset=0)`.
  2. Apply filtering logic on `user_id`, `flight_id`, `seat_class`, and `mismatch_only`.
  3. Return a dict containing `summary` (mismatch counts and totals across the matching scope) and `items` (paginated results).
  4. Provide clear, descriptive docstrings explaining tool parameters and output structure.
- **Relevant Context**:
  - `backoffice/server.py`: `bookings_search()`, `bookings_search_quotes()`.
- **Status**: `[ ] pending`

---

### Sub-Task 3: Verification & Integration Testing
- **Intent**: Verify that the tool registers correctly in the MCP server, responds properly with snapshot data, and respects filtering/pagination.
- **Expected Outcomes**:
  - Tool appears in MCP tools list under the `bookings` system.
  - Queries return expected mismatch numbers matching the snapshot data (e.g., October 2099 test cases).
  - Filtering by `user_id`, `flight_id`, `seat_class`, and `mismatch_only=False` behaves as expected.
- **Todo List**:
  1. Run Python validation / syntax check on `backoffice/server.py`.
  2. Execute a quick verification script / MCP call against the snapshot data to test default and filtered outputs.
  3. Confirm error handling and edge cases (e.g., empty filters, no matches).
- **Relevant Context**:
  - `snapshot/bookings.csv`, `snapshot/quotes.csv`, `snapshot/flights.csv`.
- **Status**: `[ ] pending`
