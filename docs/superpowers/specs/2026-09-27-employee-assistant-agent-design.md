# Employee Assistant Agent MVP Design

Date: 2026-09-27
Status: Draft for user review
Project: Internal employee assistant agent built with Pi Agent SDK

## 1. Problem

Employees often need to coordinate one office task across several internal systems:

- Check a product/material state in Teamcenter.
- Create or prepare an OA workflow.
- Notify stakeholders through Tencent RTX, an older internal instant messaging tool.
- Schedule a review meeting in Outlook.

The goal of this MVP is not to directly integrate production enterprise systems yet. The goal is to build a runnable command-line Agent that proves the full execution loop:

`user request -> Agent state -> model decision -> tool call -> observation -> confirmation -> side effect -> trace -> final answer`

This makes the project useful both as an engineering prototype and as a portfolio/interview project.

## 2. MVP Scenario

The first supported scenario is:

> "Help me check material P-10086 in Teamcenter. If it is released, create an OA change review draft, notify Zhang San and Li Si in Tencent RTX, and schedule a review meeting tomorrow afternoon in Outlook."

Expected behavior:

1. The CLI receives the user's natural language request.
2. The Agent extracts the material number, task goal, stakeholders, and meeting intent.
3. The Agent queries the Teamcenter mock adapter.
4. If the material is released, the Agent prepares an OA workflow draft.
5. Before any side-effecting operation, the runtime pauses and asks for human confirmation.
6. After confirmation, the Agent sends a Tencent RTX mock notification and creates an Outlook mock meeting.
7. The Agent writes an append-only JSONL trace for every material event.
8. The CLI prints a final summary with receipts and the trace file path.

## 3. Non-Goals

This MVP will not:

- Connect to production Teamcenter, OA, Tencent RTX, or Outlook.
- Implement a web UI.
- Implement multi-agent coordination.
- Store enterprise secrets.
- Send real messages, create real OA workflows, or create real Outlook meetings.
- Depend on any specific Tencent RTX API being available.

## 4. Recommended Approach

Use a single Agent core with typed tool ports, mock adapters, a Pi SDK model adapter, and an explicit runtime confirmation gate.

This is preferred over:

- Rule-only orchestration, which is safer but does not demonstrate an Agent decision loop well.
- Multi-agent collaboration, which increases state, permission, and audit complexity before the single-Agent MVP proves real value.

## 5. Architecture

```text
CLI
  -> Agent Runtime
  -> Pi SDK Model Adapter
  -> Agent Core
  -> Tool Ports
       -> Teamcenter Mock Adapter
       -> OA Mock Adapter
       -> Tencent RTX Mock Adapter
       -> Outlook Mock Adapter
  -> JSONL Event Log
  -> CLI Final Output
```

## 6. Component Responsibilities

### CLI

The CLI is the initial product surface.

Responsibilities:

- Accept the user's request as a command argument or prompt input.
- Start one Agent run.
- Render planned actions and confirmation prompts.
- Print the final answer.
- Print the trace file location.

Example:

```bash
npm run dev -- "Check material P-10086 and arrange a review"
```

### Agent Core

The Agent Core owns domain-neutral Agent behavior and must not import Pi SDK types directly.

Responsibilities:

- Maintain `AgentState`.
- Interpret model decisions.
- Validate tool intents.
- Apply observations to state.
- Decide when a run is complete, blocked, failed, or waiting for confirmation.

Core concepts:

- `AgentRun`
- `AgentState`
- `Decision`
- `ToolIntent`
- `Observation`
- `AgentResult`
- `ConfirmationRequest`

### Pi SDK Model Adapter

The Pi adapter sits behind a model port.

Responsibilities:

- Call Pi Agent SDK.
- Convert runtime state and recent events into the model prompt/input.
- Convert the model result into the project's `Decision` type.
- Keep SDK-specific types out of core logic.

Before implementing the real adapter, the repository must include an adapter mapping note explaining:

- Which Pi SDK APIs are used.
- How Pi SDK output maps to `Decision`.
- How malformed model output is handled.
- How cancellation, budget, and trace metadata are propagated.

### Runtime

The runtime executes the Agent loop.

Responsibilities:

- Create a run id.
- Enforce step budgets.
- Enforce cancellation points.
- Execute tool calls.
- Require confirmation for side effects.
- Write append-only JSONL events.
- Redact sensitive data before logging.
- Return a structured final result.

The runtime should preserve the loop:

```text
state -> model decision -> validated tool intent -> tool call -> observation -> state update
```

### Tool Ports

Each enterprise system is represented by a port interface.

Tool ports:

- `TeamcenterPort`
- `OaPort`
- `TencentRtxPort`
- `OutlookPort`

Adapters for this MVP are mocks with deterministic sample data.

## 7. Tool Contracts

### Teamcenter

Read-only in MVP.

Capabilities:

- `getMaterial(materialNo)`

Sample output:

```json
{
  "materialNo": "P-10086",
  "name": "Hydraulic clamp assembly",
  "revision": "B",
  "lifecycleState": "Released",
  "owner": "zhang.san",
  "lastModifiedAt": "2026-09-26T10:30:00+08:00"
}
```

### OA

Side-effecting operations must be confirmed before execution.

Capabilities:

- `createChangeReviewDraft(input)`

MVP behavior:

- Creates a mock draft receipt.
- Does not submit a real workflow.
- Returns a stable mock `draftId`.

### Tencent RTX

Tencent RTX is treated as a legacy communication adapter.

Capabilities:

- `sendNotification(input)`

MVP behavior:

- Records a mock message receipt.
- Preserves adapter boundaries for future RTX SDK, COM, HTTP gateway, or internal bridge implementation.

### Outlook

Side-effecting operations must be confirmed before execution.

Capabilities:

- `createMeeting(input)`

MVP behavior:

- Creates a mock meeting receipt.
- Uses an interface compatible with a future Microsoft Graph implementation.

## 8. Security and Policy

Default security posture:

- Unknown tools are denied.
- Side-effecting operations require explicit confirmation.
- Tool inputs are validated before execution.
- Tool outputs are treated as untrusted.
- Trace logging redacts secrets and tokens.
- The Agent cannot execute arbitrary shell commands.
- The Agent cannot call connectors that are not registered in the allowlist.

Side-effecting tools:

- OA draft creation.
- Tencent RTX notification.
- Outlook meeting creation.

Read-only tools:

- Teamcenter material lookup.

## 9. Event Log

The event log is append-only JSONL.

Each event should include:

- `runId`
- `timestamp`
- `type`
- `step`
- `payload`
- `traceId`

Important event types:

- `run.started`
- `model.decision.requested`
- `model.decision.received`
- `tool.intent.validated`
- `tool.call.started`
- `tool.call.completed`
- `confirmation.requested`
- `confirmation.received`
- `run.completed`
- `run.failed`

## 10. Failure Behavior

The MVP should handle:

- Missing material number.
- Material not found.
- Material not released.
- Model returns malformed decision.
- Tool adapter throws an error.
- Confirmation is denied.
- Step budget is exhausted.

The Agent should return a clear final state instead of silently failing.

## 11. Test Strategy

Testing should be built in this order:

1. Pure core tests with fake model decisions.
2. Tool port contract tests for mock adapters.
3. Runtime loop tests covering success, rejection, and failure paths.
4. Golden run tests for representative CLI scenarios.
5. Trace tests verifying event ordering and redaction.

## 12. Implementation Phases

### Phase 1: Project Skeleton

- TypeScript project setup.
- Basic scripts.
- Source layout.
- Shared domain types.

### Phase 2: Pure Core

- Define Agent state and decision types.
- Implement state transitions.
- Implement tool intent validation.

### Phase 3: Mock Adapters

- Teamcenter mock.
- OA mock.
- Tencent RTX mock.
- Outlook mock.

### Phase 4: Runtime

- Agent loop.
- Step budget.
- Confirmation gate.
- JSONL event log.
- Final result rendering.

### Phase 5: Pi SDK Adapter

- Add `@earendil-works/pi-coding-agent@0.84.2`.
- Implement model port adapter.
- Add mapping note.
- Keep SDK types outside core.

### Phase 6: CLI Golden Runs

- Add runnable demo commands.
- Add README scenario.
- Add trace examples.

## 13. Acceptance Criteria

The MVP is complete when:

- A user can run the CLI with the sample business request.
- The Agent queries Teamcenter mock data.
- The Agent prepares an OA draft only when the material is released.
- The runtime pauses before side effects.
- After confirmation, the Agent records mock RTX and Outlook receipts.
- The Agent writes a readable JSONL trace.
- Core logic can be tested without Pi SDK or enterprise systems.
- Tool adapters are replaceable.
- The final CLI output explains what happened and where the trace is stored.

## 14. Open Questions

These are intentionally deferred beyond MVP:

- Which exact Teamcenter interface will be used in production: REST, SOA, ITK wrapper, or internal service?
- Which OA platform and authentication method will be used?
- Does Tencent RTX expose a stable SDK, COM automation, database gateway, or internal message bridge?
- Will Outlook integration use Microsoft Graph delegated auth or application permissions?
- Which identity provider maps employees across Teamcenter, OA, RTX, and Outlook?

