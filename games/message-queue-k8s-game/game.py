#!/usr/bin/env python3
"""Cluster Commander — a terminal game for learning message queues & Kubernetes.

You run a message-processing platform. Producers fire messages into a QUEUE; you
scale consumer PODS across cluster NODES to drain it before it overflows. Survive
escalating traffic waves while keeping your SLA high and your budget positive.

Pure standard library — no installs. Run it:

    python3 game.py                 # interactive play
    python3 game.py --demo          # auto-pilot a few ticks, then exit (smoke test)
    python3 game.py --seed 1        # reproducible run

Concepts taught (via lesson popups that fire when each first becomes relevant):
queues/producers/consumers, backpressure, deployments & replicas, scheduling /
pending pods, self-healing, horizontal pod autoscaling, acks & data loss.
"""

from __future__ import annotations

import argparse
import os
import random
import sys
import textwrap
from dataclasses import dataclass, field

from lessons import LESSONS

# ---------------------------------------------------------------------------
# Tuning constants
# ---------------------------------------------------------------------------

QUEUE_CAPACITY = 200          # messages the broker topic can buffer before dropping
POD_THROUGHPUT = 6            # messages a healthy pod drains per tick
PODS_PER_NODE = 3             # how many pods a single node can host (resource limits)
POD_COST = 2                  # budget spent per running pod per tick
NODE_COST = 40                # one-off budget cost to add a node
START_BUDGET = 400
START_NODES = 1
CRASH_CHANCE = 0.04           # per-pod chance to crash on a tick
RESTART_DELAY = 2             # ticks before Kubernetes self-heals a crashed pod
HPA_STEP = 1                  # max replica change per tick when autoscaling
WAVES = 5                     # waves to clear to win
TICKS_PER_WAVE = 12
SLA_FAIL_THRESHOLD = 80.0     # SLA % below which you lose

POD_RUNNING = "running"
POD_PENDING = "pending"       # wants to run but no node capacity (scheduling)
POD_CRASHED = "crashed"       # down; will be self-healed after RESTART_DELAY ticks


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Pod:
    state: str = POD_PENDING
    restart_in: int = 0       # ticks until a crashed pod is restarted


@dataclass
class Game:
    rng: random.Random
    queue: int = 0
    desired_replicas: int = 1
    nodes: int = START_NODES
    budget: int = START_BUDGET
    pods: list[Pod] = field(default_factory=list)
    hpa_on: bool = False
    hpa_target: int = 40              # target queue depth the HPA aims to hold
    wave: int = 1
    tick_in_wave: int = 0
    produced_total: int = 0
    processed_total: int = 0
    dropped_total: int = 0
    fired_lessons: set[str] = field(default_factory=set)
    pending_lessons: list[str] = field(default_factory=list)
    game_over: bool = False
    won: bool = False

    # -- capacity helpers ---------------------------------------------------

    @property
    def node_pod_capacity(self) -> int:
        return self.nodes * PODS_PER_NODE

    @property
    def running_pods(self) -> int:
        return sum(1 for p in self.pods if p.state == POD_RUNNING)

    @property
    def pending_pods(self) -> int:
        return sum(1 for p in self.pods if p.state == POD_PENDING)

    @property
    def crashed_pods(self) -> int:
        return sum(1 for p in self.pods if p.state == POD_CRASHED)

    @property
    def sla(self) -> float:
        # Share of produced messages that were NOT dropped.
        if self.produced_total == 0:
            return 100.0
        return 100.0 * (1 - self.dropped_total / self.produced_total)

    def producer_rate(self) -> int:
        # Traffic ramps up each wave, with a little jitter.
        base = 5 + (self.wave - 1) * 5
        return base + self.rng.randint(0, 3)

    # -- lessons ------------------------------------------------------------

    def queue_lesson(self, event: str) -> None:
        """Queue a lesson to be shown once, the first time its event happens."""
        if event in LESSONS and event not in self.fired_lessons:
            self.fired_lessons.add(event)
            self.pending_lessons.append(event)


# ---------------------------------------------------------------------------
# Simulation step
# ---------------------------------------------------------------------------

def reconcile_pods(g: Game) -> None:
    """Make the pod list match desired_replicas, respecting node capacity.

    This is the Kubernetes control loop in miniature: add/remove pods toward the
    desired count, and leave pods PENDING when no node can host them.
    """
    # Remove surplus pods (prefer removing pending, then crashed, then running).
    while len(g.pods) > g.desired_replicas:
        for state in (POD_PENDING, POD_CRASHED, POD_RUNNING):
            victim = next((p for p in g.pods if p.state == state), None)
            if victim:
                g.pods.remove(victim)
                break
    # Add missing pods (they start pending until scheduled).
    while len(g.pods) < g.desired_replicas:
        g.pods.append(Pod(state=POD_PENDING))

    # Schedule pending pods onto available node capacity.
    schedulable = max(0, g.node_pod_capacity - g.running_pods)
    for pod in g.pods:
        if pod.state == POD_PENDING and schedulable > 0:
            pod.state = POD_RUNNING
            schedulable -= 1

    if g.pending_pods > 0:
        g.queue_lesson("scheduling")


def run_hpa(g: Game) -> None:
    """Horizontal Pod Autoscaler: nudge desired_replicas toward the queue target."""
    if not g.hpa_on:
        return
    if g.queue > g.hpa_target and g.desired_replicas < g.node_pod_capacity:
        g.desired_replicas = min(g.desired_replicas + HPA_STEP, g.node_pod_capacity)
        if g.desired_replicas > 1:
            g.queue_lesson("langgraph")
    elif g.queue < g.hpa_target // 2 and g.desired_replicas > 1:
        g.desired_replicas = max(1, g.desired_replicas - HPA_STEP)


def step(g: Game) -> list[str]:
    """Advance the simulation by one tick. Returns a list of event log lines."""
    log: list[str] = []

    run_hpa(g)
    reconcile_pods(g)

    # 1. Produce messages into the queue (drop on overflow).
    incoming = g.producer_rate()
    g.produced_total += incoming
    g.queue_lesson("queue_intro")
    space = QUEUE_CAPACITY - g.queue
    accepted = min(incoming, space)
    dropped = incoming - accepted
    g.queue += accepted
    if dropped > 0:
        g.dropped_total += dropped
        log.append(f"! Queue full — dropped {dropped} message(s)")
        g.queue_lesson("dropped")
    log.append(f"+ Producers sent {incoming} msg (wave {g.wave})")

    # 2. Consume: each running pod drains up to POD_THROUGHPUT messages.
    capacity = g.running_pods * POD_THROUGHPUT
    processed = min(capacity, g.queue)
    g.queue -= processed
    g.processed_total += processed
    log.append(f"- {g.running_pods} pod(s) processed {processed} msg")
    if processed > 0:
        # Each processed message is one LangChain invocation run by the worker.
        g.queue_lesson("langchain")

    # Backpressure: backlog growing despite consuming.
    if accepted > capacity and g.queue > g.hpa_target:
        g.queue_lesson("backpressure")

    # 3. Random pod crashes (design for failure).
    for pod in g.pods:
        if pod.state == POD_RUNNING and g.rng.random() < CRASH_CHANCE:
            pod.state = POD_CRASHED
            pod.restart_in = RESTART_DELAY
            log.append("x A pod crashed")
            g.queue_lesson("self_healing")

    # 4. Self-heal crashed pods after their restart delay.
    for pod in g.pods:
        if pod.state == POD_CRASHED:
            pod.restart_in -= 1
            if pod.restart_in <= 0:
                # Restart, subject to node capacity (else back to pending).
                if g.running_pods < g.node_pod_capacity:
                    pod.state = POD_RUNNING
                    log.append("✓ Kubernetes restarted a pod")
                else:
                    pod.state = POD_PENDING

    # 5. Charge budget for running pods.
    g.budget -= g.running_pods * POD_COST

    # 6. Advance wave / tick clock.
    g.tick_in_wave += 1
    if g.tick_in_wave >= TICKS_PER_WAVE:
        g.tick_in_wave = 0
        if g.wave >= WAVES:
            g.game_over = True
            g.won = g.sla >= SLA_FAIL_THRESHOLD
        else:
            g.wave += 1
            log.append(f"=== WAVE {g.wave} — traffic rising ===")
            g.queue_lesson("wave")

    # 7. Loss conditions.
    if g.sla < SLA_FAIL_THRESHOLD and g.produced_total > 30:
        g.game_over = True
        g.won = False
    if g.budget < 0:
        g.game_over = True
        g.won = False

    return log


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def clear_screen() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def bar(value: int, total: int, width: int = 30) -> str:
    total = max(total, 1)
    filled = int(width * min(value, total) / total)
    return "[" + "#" * filled + "-" * (width - filled) + "]"


def box(title: str, body: str, width: int = 64) -> str:
    inner = width - 2
    lines = ["+" + "-" * inner + "+"]
    lines.append("|" + title.center(inner) + "|")
    lines.append("+" + "-" * inner + "+")
    for para in body.split("\n"):
        for wrapped in textwrap.wrap(para, inner - 2) or [""]:
            lines.append("| " + wrapped.ljust(inner - 2) + " |")
    lines.append("+" + "-" * inner + "+")
    return "\n".join(lines)


def render(g: Game, log: list[str]) -> None:
    clear_screen()
    print("=" * 64)
    print("  CLUSTER COMMANDER  —  Message Queue & Kubernetes Simulator")
    print("  (pods = LangGraph agent workers · messages = LangChain tasks)")
    print("=" * 64)
    print(f"  Wave {g.wave}/{WAVES}   tick {g.tick_in_wave}/{TICKS_PER_WAVE}"
          f"   Budget ${g.budget}   SLA {g.sla:5.1f}%")
    print()
    print(f"  Queue  {bar(g.queue, QUEUE_CAPACITY)} {g.queue}/{QUEUE_CAPACITY}")
    print(f"  Pods   running={g.running_pods}  pending={g.pending_pods}"
          f"  crashed={g.crashed_pods}  desired={g.desired_replicas}")
    print(f"  Nodes  {g.nodes}  (capacity {g.node_pod_capacity} pods)"
          f"   HPA={'ON' if g.hpa_on else 'off'} (target depth {g.hpa_target})")
    print(f"  Totals produced={g.produced_total}  processed={g.processed_total}"
          f"  dropped={g.dropped_total}")
    print("-" * 64)
    if log:
        for line in log[-7:]:
            print("  " + line)
        print("-" * 64)


def show_lessons(g: Game, interactive: bool) -> None:
    while g.pending_lessons:
        event = g.pending_lessons.pop(0)
        title, body = LESSONS[event]
        print()
        print(box(f"LESSON · {title}", body))
        if interactive:
            input("\n  (press Enter to continue) ")


# ---------------------------------------------------------------------------
# Command handling
# ---------------------------------------------------------------------------

HELP = """\
Commands:
  tick [n]      advance the simulation n ticks (default 1)
  scale <n>     set desired pod replicas to n
  up / down     +1 / -1 desired replicas
  hpa [target]  toggle autoscaling on/off; optional queue-depth target
  node          add a node (+capacity) for $%d
  help          show this help
  quit          leave the game
""" % NODE_COST


def handle_command(g: Game, raw: str) -> tuple[bool, int]:
    """Process one command. Returns (keep_playing, ticks_to_run)."""
    parts = raw.strip().split()
    if not parts:
        return True, 0
    cmd, *args = parts

    if cmd in ("quit", "q", "exit"):
        return False, 0
    if cmd in ("help", "h", "?"):
        print(HELP)
        input("  (press Enter) ")
        return True, 0
    if cmd in ("tick", "t", "run", ""):
        n = int(args[0]) if args and args[0].isdigit() else 1
        return True, max(1, n)
    if cmd == "scale" and args and args[0].isdigit():
        g.desired_replicas = max(0, int(args[0]))
        g.queue_lesson("replicas")
        if g.desired_replicas > 1:
            g.queue_lesson("langgraph")
        return True, 0
    if cmd == "up":
        g.desired_replicas += 1
        g.queue_lesson("replicas")
        if g.desired_replicas > 1:
            g.queue_lesson("langgraph")
        return True, 0
    if cmd == "down":
        g.desired_replicas = max(0, g.desired_replicas - 1)
        return True, 0
    if cmd == "hpa":
        g.hpa_on = not g.hpa_on
        if args and args[0].isdigit():
            g.hpa_target = int(args[0])
        if g.hpa_on:
            g.queue_lesson("hpa")
        return True, 0
    if cmd == "node":
        if g.budget >= NODE_COST:
            g.budget -= NODE_COST
            g.nodes += 1
        else:
            print("  Not enough budget for a node.")
            input("  (press Enter) ")
        return True, 0

    print(f"  Unknown command: {raw!r}. Type 'help'.")
    input("  (press Enter) ")
    return True, 0


# ---------------------------------------------------------------------------
# Game loops
# ---------------------------------------------------------------------------

def play_interactive(g: Game) -> None:
    log: list[str] = []
    show_lessons(g, interactive=True)
    while not g.game_over:
        render(g, log)
        show_lessons(g, interactive=True)
        try:
            raw = input("\n  command (tick/scale/hpa/node/up/down/help/quit) > ")
        except (EOFError, KeyboardInterrupt):
            print("\n  Goodbye.")
            return
        keep, ticks = handle_command(g, raw)
        if not keep:
            print("  Goodbye.")
            return
        log = []
        for _ in range(ticks):
            if g.game_over:
                break
            log = step(g)
            show_lessons(g, interactive=True)
    render(g, log)
    finish(g)


def play_demo(g: Game, ticks: int = 18) -> None:
    """Auto-pilot run for smoke testing (non-interactive)."""
    print("Running auto-pilot demo...\n")
    g.hpa_on = True
    g.queue_lesson("hpa")
    show_lessons(g, interactive=False)
    for i in range(ticks):
        if g.game_over:
            break
        # Naive auto-pilot: add a node if pods are stuck pending and we can afford it.
        if g.pending_pods > 0 and g.budget >= NODE_COST:
            g.budget -= NODE_COST
            g.nodes += 1
        log = step(g)
        show_lessons(g, interactive=False)
        print(f"[tick {i+1}] queue={g.queue:3d} running={g.running_pods} "
              f"pending={g.pending_pods} sla={g.sla:5.1f}% budget=${g.budget}")
        for line in log:
            print("    " + line)
    print()
    if g.game_over:
        finish(g)
    else:
        print(box("DEMO COMPLETE",
                  f"Auto-pilot ran {ticks} ticks and reached wave {g.wave} with "
                  f"an SLA of {g.sla:.1f}% (dropped {g.dropped_total}). The full "
                  f"game runs until all {WAVES} waves are cleared — play it with "
                  f"`python3 game.py`."))


def finish(g: Game) -> None:
    print()
    if g.won:
        print(box("YOU WIN!",
                  f"You cleared all {WAVES} waves with an SLA of {g.sla:.1f}%. "
                  f"Your cluster scaled with demand and kept the queue drained. "
                  f"Final budget: ${g.budget}."))
    else:
        reason = ("your SLA fell below the threshold" if g.sla < SLA_FAIL_THRESHOLD
                  else "you ran out of budget" if g.budget < 0
                  else "the run ended")
        print(box("GAME OVER",
                  f"Run ended because {reason}. SLA {g.sla:.1f}%, "
                  f"dropped {g.dropped_total} messages. Tip: scale pods (or enable "
                  f"HPA) before the backlog hits capacity, and add nodes so pods "
                  f"aren't stuck Pending."))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Cluster Commander game.")
    parser.add_argument("--demo", action="store_true",
                        help="auto-pilot a few ticks then exit (smoke test)")
    parser.add_argument("--seed", type=int, default=None,
                        help="seed the RNG for reproducible runs")
    parser.add_argument("--ticks", type=int, default=18,
                        help="number of ticks for --demo")
    args = parser.parse_args(argv)

    rng = random.Random(args.seed)
    g = Game(rng=rng, desired_replicas=1)
    g.queue_lesson("course_intro")  # framing: how this ties back to the course

    if args.demo:
        play_demo(g, ticks=args.ticks)
    else:
        play_interactive(g)
    return 0


if __name__ == "__main__":
    sys.exit(main())
