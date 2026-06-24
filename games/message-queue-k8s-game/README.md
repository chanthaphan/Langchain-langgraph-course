# Cluster Commander 🎮

A terminal game for learning **message queues**, **Kubernetes**, and how they tie
back to **LangChain / LangGraph** — by running a system instead of reading about one.

You command a message-processing platform. Producers fire **messages** (agent tasks)
into a **queue**; you scale **consumer pods** — each one a LangGraph agent worker
running LangChain chains — across cluster **nodes** to drain the queue before it
overflows. Survive escalating traffic waves while keeping your SLA high and your
budget positive.

It's pure Python standard library — **nothing to install.**

## Run it

```bash
cd games/message-queue-k8s-game
python3 game.py            # interactive play
python3 game.py --demo     # auto-pilot a few ticks (smoke test / quick look)
python3 game.py --seed 1   # reproducible run
```

## How to play

Each turn you see a live dashboard (queue depth, pods, nodes, SLA, budget) and type
a command. Then you advance the simulation and watch what happens.

| Command        | What it does                                                  |
| -------------- | ------------------------------------------------------------ |
| `tick [n]`     | Advance the simulation by `n` ticks (default 1)              |
| `scale <n>`    | Set desired pod **replicas** to `n`                          |
| `up` / `down`  | +1 / −1 desired replicas                                     |
| `hpa [target]` | Toggle **autoscaling** on/off; optional queue-depth target   |
| `node`         | Add a **node** (more pod capacity) for budget                |
| `help`         | Show the command list                                        |
| `quit`         | Leave the game                                               |

**Goal:** clear all waves with your SLA above the threshold. **You lose** if the SLA
collapses (too many dropped messages) or you run out of budget.

**Core tension:** more pods = faster draining but higher cost; pods need node
capacity or they sit `Pending`; pods crash and self-heal; traffic keeps rising.
Scale ahead of the curve.

## What you'll learn

Lessons pop up the first time each concept becomes relevant, so the explanation
arrives exactly when it matters:

**Message queues** — producers, consumers, broker/topic decoupling · backpressure ·
queue capacity, acknowledgements & data loss on overflow.

**Kubernetes** — Deployments & replicas · pods, nodes & scheduling (`Pending` pods
when capacity runs out) · self-healing of crashed pods · horizontal pod autoscaling
(HPA) driven by queue depth.

**LangChain / LangGraph** — why production LLM agents run as queued, containerized
workers: each message is a LangChain chain invocation (slow, rate-limited LLM calls
are why you queue), and each worker pod is a stateful LangGraph agent graph whose
checkpointing pairs naturally with Kubernetes self-healing. Scaling replicas =
scaling your agents horizontally.

## Files

- `game.py` — simulation engine, dashboard, and command loop
- `lessons.py` — the lesson content shown during play

## Note

The simulation is a teaching abstraction, not a faithful Kafka/Kubernetes emulator —
the mechanics are simplified so the *concepts* come through clearly.
