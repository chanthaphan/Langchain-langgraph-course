---
name: loop-me
description: |
  Run a prompt or slash command on a recurring interval until told to stop.
  Use when the user wants to set up a recurring task, poll for status, watch
  for a condition, or run something repeatedly on an interval.

  Trigger for:
  - "loop me <task>", "/loop-me 5m <task>", "keep running X every N minutes"
  - "poll the build every 5 minutes", "check the deploy until it's green"
  - "watch for X and re-run Y", "repeat this until <condition>"

  Don't trigger for:
  - One-off tasks that should run exactly once
  - Long-running blocking processes (use a background command instead)
---

# Loop Me

Run a prompt or slash command repeatedly on a fixed interval, re-invoking
yourself each cycle, until a stop condition is met or the user cancels.

## Arguments

Invoked as `/loop-me [interval] <prompt-or-command>`.

- `interval` — optional. A duration like `30s`, `5m`, `1h`. Defaults to `10m`.
- `prompt-or-command` — required. The work to run each cycle. May be plain
  text or a slash command (e.g. `/code-review`).

Examples:
- `/loop-me 5m check CI status and report failures`
- `/loop-me /run-tests`
- `/loop-me 1h summarize new commits on main`

## How to run the loop

1. **Parse** the interval and the task. If no interval is given, use `10m`.
   Confirm to the user what will run and how often.

2. **Run one iteration now.** Execute the task exactly as if the user had
   typed the prompt or slash command directly. Report the result concisely.

3. **Check the stop condition.** Stop looping when any of these is true:
   - The user asked to loop "until X" and X is now satisfied.
   - The user said to stop / cancel.
   - The task has failed in a way that repeating cannot fix (report and stop).

4. **Schedule the next iteration.** If not stopping, use `ScheduleWakeup`
   (dynamic `/loop` pacing) to re-fire this same `/loop-me` invocation after
   the interval. Pass the identical input back so the next wake-up repeats the
   task. Choose `delaySeconds` from the interval, honoring cache-window
   guidance (stay under 300s when actively polling; use 1200s+ for idle ticks).

5. **End the turn** after scheduling. Do not block with `sleep`. When the
   wake-up fires, resume at step 2.

## Stopping

Omit the `ScheduleWakeup` call to end the loop. Always stop immediately when
the user says stop, and confirm the loop has ended.

## Notes

- Each iteration is independent — keep per-cycle output short so a long-running
  loop does not flood the conversation.
- If an iteration errors transiently (network, rate limit), keep looping; if it
  errors structurally (bad command, missing file), stop and report.
