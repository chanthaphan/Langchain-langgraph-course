# Langchain-langgraph-course

## Modules

### 🎮 Cluster Commander — Message Queue & Kubernetes game

A pure-Python terminal game for learning **message queues**, **Kubernetes**, and how
they tie back to **LangChain / LangGraph** — by running a system, not just reading
about one. You scale consumer pods (LangGraph agent workers) on a cluster to drain a
message queue through escalating traffic waves, with lessons that pop up as each
concept first becomes relevant.

```bash
cd games/message-queue-k8s-game
python3 game.py            # play
python3 game.py --demo     # quick auto-pilot look
```

See [`games/message-queue-k8s-game/README.md`](games/message-queue-k8s-game/README.md)
for the full guide.
