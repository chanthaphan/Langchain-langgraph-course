"""Lesson content for Cluster Commander.

Each lesson is keyed by an event id. The game's LessonManager fires a lesson the
first time its event happens, so the explanation arrives exactly when the concept
becomes relevant. Keep each lesson short — a title and a few wrapped lines.
"""

LESSONS = {
    "course_intro": (
        "WELCOME — WHY THIS GAME IS IN A LANGCHAIN/LANGGRAPH COURSE",
        "You're building LLM apps with LangChain & LangGraph. In production those "
        "agents don't run one request at a time in a notebook — they run as many "
        "containerized WORKERS that pull tasks from a message QUEUE and execute "
        "on a Kubernetes cluster. In this game the messages are agent tasks, and "
        "each consumer POD is a LangGraph worker. Master queues + Kubernetes here "
        "and you'll know how to actually SHIP what you build in this course.",
    ),
    "queue_intro": (
        "MESSAGE QUEUES 101",
        "Producers create work (messages) and drop it onto a QUEUE (a broker "
        "TOPIC). They do NOT call consumers directly. Consumers pull messages "
        "off the queue and process them at their own pace. This decoupling lets "
        "each side scale and fail independently — the queue is a shock absorber "
        "between them.",
    ),
    "backpressure": (
        "BACKPRESSURE",
        "Messages are arriving faster than your consumers can process them, so "
        "the backlog is growing. That gap is backpressure. You have two levers: "
        "process faster (add consumer pods) or slow producers down. If you do "
        "neither, the queue fills up.",
    ),
    "replicas": (
        "DEPLOYMENTS & REPLICAS",
        "In Kubernetes you don't start consumers by hand. A Deployment declares "
        "a desired REPLICA count, and Kubernetes keeps exactly that many PODS "
        "running. Scaling = changing one number; the control loop does the rest. "
        "More replicas = more total throughput draining the queue.",
    ),
    "scheduling": (
        "SCHEDULING & PENDING PODS",
        "Pods run on NODES, and every node has finite CPU/memory. When you ask "
        "for more pods than your nodes can host, the scheduler can't place them "
        "and they sit in PENDING — counting for nothing. Add a node (capacity) "
        "to unblock them. This is why a cluster can't scale forever for free.",
    ),
    "self_healing": (
        "SELF-HEALING",
        "A pod just crashed. Because a Deployment guarantees a desired replica "
        "count, Kubernetes notices the missing pod and automatically restarts a "
        "replacement after a short delay. You design for failure instead of "
        "preventing it — individual pods are cattle, not pets.",
    ),
    "hpa": (
        "HORIZONTAL POD AUTOSCALING (HPA)",
        "Instead of scaling pods by hand, the HPA watches a metric and adjusts "
        "replicas automatically. Here it scales on QUEUE DEPTH: backlog above "
        "the target adds pods, a drained queue removes them. Autoscaling chases "
        "demand so you don't have to babysit every traffic spike.",
    ),
    "dropped": (
        "ACKNOWLEDGEMENTS & DATA LOSS",
        "The queue hit its capacity and a message was DROPPED — that's lost "
        "work and a hit to your SLA. Real brokers bound how much they buffer; "
        "once full, new messages are rejected. Consumers ACK a message only "
        "after processing it, so nothing is lost mid-flight — but an overflowing "
        "queue still sheds load. Keep the backlog below capacity.",
    ),
    "langchain": (
        "LANGCHAIN — WHAT EACH WORKER ACTUALLY DOES",
        "Every message a pod processes is one LangChain invocation: a CHAIN of "
        "prompt -> LLM -> output parser (often with tools or retrieval). Those "
        "LLM calls are slow and rate-limited — seconds each, not microseconds. "
        "That latency is exactly why you put a queue in front: producers enqueue "
        "tasks instantly, and a pool of workers drains them as fast as the LLM "
        "allows. Throughput = (workers) x (chains each can run per tick).",
    ),
    "langgraph": (
        "LANGGRAPH — THE WORKER IS A STATEFUL AGENT GRAPH",
        "Each pod runs a LangGraph app: a graph of nodes with shared state, "
        "loops, and tool calls — a real agent, not a single prompt. LangGraph's "
        "checkpointing means a task can be resumed if a worker dies mid-run, "
        "which pairs perfectly with Kubernetes self-healing. Scaling replicas "
        "runs MORE copies of the same graph in parallel — horizontal scaling of "
        "your agents. That's how a LangGraph app serves production traffic.",
    ),
    "wave": (
        "TRAFFIC WAVES",
        "Production traffic isn't steady — it comes in waves. Each wave ramps "
        "the producer rate higher. Capacity that was comfortable a minute ago "
        "becomes backpressure now. Watch the trend, not just the current depth, "
        "and scale ahead of the curve.",
    ),
}
