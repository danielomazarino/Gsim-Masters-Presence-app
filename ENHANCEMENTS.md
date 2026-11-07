# Träningsnärvaro (training presence app) — Enhancements & Next Development Phase

## Current baseline

The current working deployment is:

Version 11
Deployed: 5 September 2026, 01:56

Version 11 is the current stable baseline.

The application is a Google Apps Script Web App using Google Sheets as the source of truth.

The application has been successfully opened and used on an iPhone 13 without an additional login prompt.

## Verified functionality — DO NOT REDESIGN

The following functionality has already been tested and is considered working:

- Attendance registration works.
- Entire member rows are touch/click targets.
- Present/absent state works.
- Date selection works:
  - Idag
  - Igår
  - 2 dagar sedan
- Location selection works:
  - Träningsgrupp A
  - Träningsgrupp B
- Attendance count works.
- Save to Google Sheet works.
- Missing training columns are automatically created when required.
- Administration view works.
- Existing training columns are listed correctly.
- Training-column deletion works.
- Deletion confirmation works.
- Physical deletion of the complete Google Sheet column works.
- System columns are protected.
- The previous administration UI problem showing "Index undefined" / "Kolumn ..." has been fixed.
- Version 11 is currently considered functionally working.

DO NOT use future enhancement work as an excuse to rewrite or refactor this already-working functionality.

## Remaining enhancement 1 — Mobile responsiveness / iPhone 13 UI

Priority: HIGH

The application works functionally on iPhone 13, but the horizontal layout of the attendance controls feels somewhat flimsy/troublesome.

The current top control area contains:

- Träningsgrupp A
- Träningsgrupp B
- Närvarande count
- Spara
- Administration gear

These controls compete for horizontal space on a narrow mobile viewport.

Goal:

Improve the mobile responsiveness and visual robustness of the attendance header, especially on iPhone 13-sized screens.

Important constraints:

- This is a UI/layout enhancement only.
- Do NOT change backend behaviour.
- Do NOT change Google Sheet behaviour.
- Do NOT change attendance logic.
- Do NOT change save semantics as part of this enhancement.
- Do NOT remove any existing functionality.
- Preserve the simple, fast attendance workflow.
- The application is primarily intended for quick use immediately after training on a phone.
- Avoid introducing large cards, complex navigation, unnecessary dialogs, or extra interaction steps.

The exact solution should be determined after inspecting the current UI in a realistic mobile viewport.

Recommended investigation:

- Test the application using VS Code/browser responsive tools.
- Test approximately iPhone 13 viewport dimensions.
- Inspect the horizontal control row.
- Check for squeezing, wrapping, overflow, uneven spacing, or controls becoming visually unstable.
- Consider whether the controls should use a more deliberate responsive layout at narrow widths.
- Preserve the existing visual language.

Do not implement this enhancement blindly. First inspect the current layout and identify the actual cause of the problem.

## Remaining enhancement 2 — Prevent unsaved attendance from being lost

Priority: VERY HIGH

Current problem:

A trainer can mark several members as present and then forget to press "Spara".

If the page is closed, refreshed, navigated away from, or otherwise loses the current client state before Save is pressed, the newly entered attendance can be lost.

This is a real usability risk because the attendance workflow happens immediately after training and should be as fail-safe as reasonably possible.

Goal:

Make the attendance workflow resistant to accidental loss of newly entered attendance.

Possible approaches to investigate:

1. Autosave changes to Google Sheet as the trainer marks attendance.
2. Provide a reliable unsaved-changes warning/indicator.
3. Combine automatic saving with a clear saved/unsaved status.
4. Another solution that prevents accidental loss without making the workflow slower or more complicated.

Important architectural constraint:

Google Sheet remains the source of truth.

Do NOT introduce:

- Firebase
- localStorage
- sessionStorage
- cookies as a persistence mechanism
- IndexedDB
- a separate database
- analytics/tracking infrastructure

Any persistence solution must fit the existing Google Apps Script + Google Sheet architecture.

Important UX requirement:

The trainer should not have to perform additional complicated actions.

The ideal workflow should remain:

Select date → select location → tap members → attendance is safely retained.

The existing "Spara" button may be retained if useful, but the design should eliminate the risk that a trainer loses attendance simply because they forgot to press it.

Before implementation, evaluate:

- number of server calls
- reliability of Google Apps Script calls
- what happens if the network is temporarily unavailable
- race conditions if several taps occur quickly
- what happens when switching date/location
- what happens if the trainer closes or reloads the page
- whether automatic creation of a missing training column remains safe
- whether duplicate writes can occur
- whether the existing backend validation can be reused
- whether save operations remain idempotent

Do not implement this enhancement until the approach has been clearly reasoned through.

## Suggested development order

1. Investigate and improve mobile responsiveness.
2. Design the safe-save/autosave behaviour.
3. Implement the safe-save solution carefully.
4. Re-test the complete attendance workflow.
5. Test with another trainer/device.
6. Create a new deployment version only after the changes have been verified.

## Development rule

Version 11 is the known-good baseline.

Future changes should be small and isolated.

Do not combine unrelated UI and backend changes in one step unless necessary.

After each meaningful change, verify that the previously tested attendance and administration functionality still works.

The priority is reliability and ease of use, not adding more features.

## Future validation

Before considering the next version production-ready, verify:

- iPhone 13 layout
- narrow mobile viewport behaviour
- attendance marking
- date switching
- location switching
- save behaviour
- accidental navigation/refresh behaviour
- automatic training-column creation
- administration
- training-column deletion
- Google Sheet integrity

END OF FILE
