# Employee Assistant Agent Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a runnable TypeScript CLI MVP for an internal employee assistant Agent that coordinates Teamcenter, OA, Tencent RTX, and Outlook through mock adapters, with confirmation gates and JSONL trace logging.

**Architecture:** The Agent Core is independent from Pi SDK and enterprise systems. The runtime drives `state -> model decision -> validated tool intent -> tool call -> observation -> state update`, while adapters implement typed ports for Teamcenter, OA, Tencent RTX, and Outlook. Pi SDK is isolated behind a model port and adapter mapping note.

**Tech Stack:** TypeScript, Node.js, Vitest, tsx, `@earendil-works/pi-coding-agent@0.84.2`, JSONL append-only event log.

**Spec:** `docs/superpowers/specs/2026-09-27-employee-assistant-agent-design.md`

## Global Constraints

- Use TypeScript only.
- The MVP product surface is a command-line interface.
- Do not connect to production Teamcenter, OA, Tencent RTX, or Outlook.
- Core logic must not import Pi SDK types directly.
- Side-effecting operations require explicit confirmation.
- Unknown tools are denied.
- Tool outputs are treated as untrusted.
- Event logs are append-only JSONL.
- The Pi dependency version is exactly `@earendil-works/pi-coding-agent@0.84.2`.
- Current workspace is not a Git repository; commit steps are listed but cannot run until Git is initialized or the project is moved into a Git repo.

---

## File Structure

- Create `package.json`: npm scripts and dependencies.
- Create `tsconfig.json`: TypeScript compiler config.
- Create `vitest.config.ts`: test config.
- Create `src/core/types.ts`: shared Agent, tool, and result types.
- Create `src/core/state.ts`: state transition helpers.
- Create `src/core/policy.ts`: tool allowlist and side-effect policy.
- Create `src/core/planner.ts`: deterministic MVP planner used by tests and fallback execution.
- Create `src/ports/teamcenter.ts`: Teamcenter port contract.
- Create `src/ports/oa.ts`: OA port contract.
- Create `src/ports/tencent-rtx.ts`: Tencent RTX port contract.
- Create `src/ports/outlook.ts`: Outlook port contract.
- Create `src/ports/model.ts`: model decision port contract.
- Create `src/adapters/mock/teamcenter.ts`: deterministic Teamcenter mock.
- Create `src/adapters/mock/oa.ts`: deterministic OA mock.
- Create `src/adapters/mock/tencent-rtx.ts`: deterministic Tencent RTX mock.
- Create `src/adapters/mock/outlook.ts`: deterministic Outlook mock.
- Create `src/adapters/pi/README.md`: Pi SDK adapter mapping note.
- Create `src/adapters/pi/pi-model-port.ts`: Pi adapter skeleton behind `ModelPort`.
- Create `src/runtime/event-log.ts`: append-only JSONL writer.
- Create `src/runtime/redaction.ts`: trace redaction helpers.
- Create `src/runtime/runtime.ts`: Agent execution loop.
- Create `src/cli/main.ts`: CLI entry point.
- Create `tests/core/*.test.ts`: core behavior tests.
- Create `tests/adapters/mock/*.test.ts`: mock adapter contract tests.
- Create `tests/runtime/*.test.ts`: runtime loop and trace tests.
- Create `tests/cli/*.test.ts`: golden CLI behavior tests.

---

### Task 1: Project Skeleton

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `vitest.config.ts`
- Create: `src/core/types.ts`
- Test: `tests/core/types.test.ts`

**Interfaces:**
- Produces: `AgentState`, `Decision`, `ToolIntent`, `Observation`, `AgentResult`, `ToolName`, `SideEffectLevel`.

- [ ] **Step 1: Write the failing test**

```ts
import { describe, expect, it } from 'vitest';
import type { ToolIntent } from '../../src/core/types';

describe('core types', () => {
  it('represents a typed Teamcenter material lookup intent', () => {
    const intent: ToolIntent = {
      id: 'intent-1',
      toolName: 'teamcenter.getMaterial',
      sideEffect: 'read',
      input: { materialNo: 'P-10086' },
    };

    expect(intent.toolName).toBe('teamcenter.getMaterial');
    expect(intent.sideEffect).toBe('read');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- tests/core/types.test.ts`
Expected: FAIL because the project and `src/core/types.ts` do not exist yet.

- [ ] **Step 3: Write minimal implementation**

Create `package.json`:

```json
{
  "name": "employee-assistant-agent",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "dev": "tsx src/cli/main.ts",
    "test": "vitest run",
    "typecheck": "tsc --noEmit"
  },
  "dependencies": {
    "@earendil-works/pi-coding-agent": "0.84.2"
  },
  "devDependencies": {
    "@types/node": "^22.0.0",
    "tsx": "^4.19.0",
    "typescript": "^5.6.0",
    "vitest": "^2.1.0"
  }
}
```

Create `tsconfig.json`:

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "strict": true,
    "esModuleInterop": true,
    "forceConsistentCasingInFileNames": true,
    "skipLibCheck": true,
    "outDir": "dist"
  },
  "include": ["src", "tests"]
}
```

Create `vitest.config.ts`:

```ts
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    environment: 'node',
  },
});
```

Create `src/core/types.ts`:

```ts
export type ToolName =
  | 'teamcenter.getMaterial'
  | 'oa.createChangeReviewDraft'
  | 'tencentRtx.sendNotification'
  | 'outlook.createMeeting';

export type SideEffectLevel = 'read' | 'write';

export interface ToolIntent {
  id: string;
  toolName: ToolName;
  sideEffect: SideEffectLevel;
  input: Record<string, unknown>;
}

export interface Decision {
  type: 'call_tool' | 'request_confirmation' | 'complete' | 'fail';
  reason: string;
  toolIntent?: ToolIntent;
  message?: string;
}

export interface Observation {
  toolName: ToolName;
  ok: boolean;
  data?: unknown;
  error?: string;
  observedAt: string;
}

export interface AgentState {
  runId: string;
  userRequest: string;
  step: number;
  observations: Observation[];
  pendingIntents: ToolIntent[];
  confirmedIntentIds: string[];
}

export interface AgentResult {
  status: 'completed' | 'failed' | 'blocked' | 'cancelled';
  summary: string;
  receipts: Record<string, unknown>[];
  tracePath: string;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test -- tests/core/types.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

Run only after the workspace is a Git repository:

```bash
git add package.json tsconfig.json vitest.config.ts src/core/types.ts tests/core/types.test.ts
git commit -m "chore: scaffold employee assistant agent"
```

---

### Task 2: Core Policy and Planner

**Files:**
- Create: `src/core/policy.ts`
- Create: `src/core/planner.ts`
- Test: `tests/core/policy.test.ts`
- Test: `tests/core/planner.test.ts`

**Interfaces:**
- Consumes: `ToolIntent`, `ToolName`, `AgentState`, `Decision`.
- Produces: `validateToolIntent(intent): PolicyResult`, `requiresConfirmation(intent): boolean`, `planNextDecision(state): Decision`.

- [ ] **Step 1: Write failing policy tests**

```ts
import { describe, expect, it } from 'vitest';
import { requiresConfirmation, validateToolIntent } from '../../src/core/policy';

describe('tool policy', () => {
  it('allows Teamcenter read-only lookup without confirmation', () => {
    const result = validateToolIntent({
      id: 'i1',
      toolName: 'teamcenter.getMaterial',
      sideEffect: 'read',
      input: { materialNo: 'P-10086' },
    });

    expect(result.ok).toBe(true);
    expect(requiresConfirmation(result.intent!)).toBe(false);
  });

  it('requires confirmation for Outlook meeting creation', () => {
    const result = validateToolIntent({
      id: 'i2',
      toolName: 'outlook.createMeeting',
      sideEffect: 'write',
      input: { subject: 'Review', attendees: ['zhang.san'] },
    });

    expect(result.ok).toBe(true);
    expect(requiresConfirmation(result.intent!)).toBe(true);
  });
});
```

- [ ] **Step 2: Write failing planner tests**

```ts
import { describe, expect, it } from 'vitest';
import { planNextDecision } from '../../src/core/planner';
import type { AgentState } from '../../src/core/types';

describe('MVP planner', () => {
  it('starts by looking up the requested material in Teamcenter', () => {
    const state: AgentState = {
      runId: 'run-1',
      userRequest: 'Check material P-10086 and arrange a review',
      step: 0,
      observations: [],
      pendingIntents: [],
      confirmedIntentIds: [],
    };

    const decision = planNextDecision(state);

    expect(decision.type).toBe('call_tool');
    expect(decision.toolIntent?.toolName).toBe('teamcenter.getMaterial');
    expect(decision.toolIntent?.input).toEqual({ materialNo: 'P-10086' });
  });
});
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `npm test -- tests/core/policy.test.ts tests/core/planner.test.ts`
Expected: FAIL because `policy.ts` and `planner.ts` do not exist.

- [ ] **Step 4: Implement policy and planner**

Create `src/core/policy.ts`:

```ts
import type { ToolIntent, ToolName } from './types';

const allowedTools = new Set<ToolName>([
  'teamcenter.getMaterial',
  'oa.createChangeReviewDraft',
  'tencentRtx.sendNotification',
  'outlook.createMeeting',
]);

export type PolicyResult =
  | { ok: true; intent: ToolIntent }
  | { ok: false; error: string };

export function validateToolIntent(intent: ToolIntent): PolicyResult {
  if (!allowedTools.has(intent.toolName)) {
    return { ok: false, error: `Tool is not allowed: ${intent.toolName}` };
  }

  if (intent.toolName === 'teamcenter.getMaterial' && intent.sideEffect !== 'read') {
    return { ok: false, error: 'Teamcenter material lookup must be read-only' };
  }

  if (intent.toolName !== 'teamcenter.getMaterial' && intent.sideEffect !== 'write') {
    return { ok: false, error: `${intent.toolName} must be marked as write` };
  }

  return { ok: true, intent };
}

export function requiresConfirmation(intent: ToolIntent): boolean {
  return intent.sideEffect === 'write';
}
```

Create `src/core/planner.ts`:

```ts
import type { AgentState, Decision, Observation } from './types';

function materialNoFrom(request: string): string | undefined {
  return request.match(/\bP-\d+\b/i)?.[0].toUpperCase();
}

function latestObservation(state: AgentState, toolName: string): Observation | undefined {
  return [...state.observations].reverse().find((item) => item.toolName === toolName);
}

export function planNextDecision(state: AgentState): Decision {
  const materialObservation = latestObservation(state, 'teamcenter.getMaterial');

  if (!materialObservation) {
    const materialNo = materialNoFrom(state.userRequest);
    if (!materialNo) {
      return { type: 'fail', reason: 'missing_material_no', message: 'No material number was found.' };
    }
    return {
      type: 'call_tool',
      reason: 'need_material_state',
      toolIntent: {
        id: `${state.runId}-teamcenter`,
        toolName: 'teamcenter.getMaterial',
        sideEffect: 'read',
        input: { materialNo },
      },
    };
  }

  if (!materialObservation.ok) {
    return { type: 'fail', reason: 'material_lookup_failed', message: String(materialObservation.error) };
  }

  const material = materialObservation.data as { lifecycleState?: string; materialNo?: string };
  if (material.lifecycleState !== 'Released') {
    return { type: 'complete', reason: 'material_not_released', message: `Material ${material.materialNo} is not released.` };
  }

  return {
    type: 'request_confirmation',
    reason: 'released_material_requires_review_actions',
    message: `Material ${material.materialNo} is released. Confirm OA, RTX, and Outlook actions.`,
  };
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test -- tests/core/policy.test.ts tests/core/planner.test.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/core/policy.ts src/core/planner.ts tests/core/policy.test.ts tests/core/planner.test.ts
git commit -m "feat: add core policy and planner"
```

---

### Task 3: Enterprise Tool Ports and Mock Adapters

**Files:**
- Create: `src/ports/teamcenter.ts`
- Create: `src/ports/oa.ts`
- Create: `src/ports/tencent-rtx.ts`
- Create: `src/ports/outlook.ts`
- Create: `src/adapters/mock/teamcenter.ts`
- Create: `src/adapters/mock/oa.ts`
- Create: `src/adapters/mock/tencent-rtx.ts`
- Create: `src/adapters/mock/outlook.ts`
- Test: `tests/adapters/mock/*.test.ts`

**Interfaces:**
- Produces: `TeamcenterPort.getMaterial`, `OaPort.createChangeReviewDraft`, `TencentRtxPort.sendNotification`, `OutlookPort.createMeeting`.

- [ ] **Step 1: Write failing adapter tests**

```ts
import { describe, expect, it } from 'vitest';
import { createMockTeamcenterPort } from '../../../src/adapters/mock/teamcenter';
import { createMockOaPort } from '../../../src/adapters/mock/oa';
import { createMockTencentRtxPort } from '../../../src/adapters/mock/tencent-rtx';
import { createMockOutlookPort } from '../../../src/adapters/mock/outlook';

describe('mock enterprise adapters', () => {
  it('returns released Teamcenter material P-10086', async () => {
    const material = await createMockTeamcenterPort().getMaterial({ materialNo: 'P-10086' });
    expect(material.lifecycleState).toBe('Released');
    expect(material.revision).toBe('B');
  });

  it('returns stable receipts for side-effecting mock tools', async () => {
    const oa = await createMockOaPort().createChangeReviewDraft({ materialNo: 'P-10086', title: 'Change Review' });
    const rtx = await createMockTencentRtxPort().sendNotification({ recipients: ['zhang.san'], message: 'Review P-10086' });
    const meeting = await createMockOutlookPort().createMeeting({ subject: 'Review P-10086', attendees: ['zhang.san'], start: '2026-09-28T14:00:00+08:00', end: '2026-09-28T15:00:00+08:00' });

    expect(oa.draftId).toBe('OA-DRAFT-P-10086');
    expect(rtx.messageId).toBe('RTX-MSG-001');
    expect(meeting.meetingId).toBe('OUTLOOK-MEETING-001');
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npm test -- tests/adapters/mock`
Expected: FAIL because port and adapter files do not exist.

- [ ] **Step 3: Implement ports and mock adapters**

Implement interfaces and deterministic mock data exactly matching the test names and return fields.

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test -- tests/adapters/mock`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/ports src/adapters/mock tests/adapters/mock
git commit -m "feat: add enterprise mock adapters"
```

---

### Task 4: Runtime, Confirmation Gate, and Event Log

**Files:**
- Create: `src/runtime/redaction.ts`
- Create: `src/runtime/event-log.ts`
- Create: `src/runtime/runtime.ts`
- Test: `tests/runtime/runtime.test.ts`
- Test: `tests/runtime/event-log.test.ts`

**Interfaces:**
- Consumes: planner, policy, tool ports.
- Produces: `runEmployeeAssistant(input): Promise<AgentResult>`, `JsonlEventLog.append(event)`.

- [ ] **Step 1: Write failing runtime success test**

```ts
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { createMockTeamcenterPort } from '../../src/adapters/mock/teamcenter';
import { createMockOaPort } from '../../src/adapters/mock/oa';
import { createMockTencentRtxPort } from '../../src/adapters/mock/tencent-rtx';
import { createMockOutlookPort } from '../../src/adapters/mock/outlook';
import { runEmployeeAssistant } from '../../src/runtime/runtime';

describe('employee assistant runtime', () => {
  it('runs the released material review flow after confirmation', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'employee-agent-'));
    const result = await runEmployeeAssistant({
      userRequest: 'Check material P-10086 and arrange a review',
      traceDir: dir,
      confirm: async () => true,
      ports: {
        teamcenter: createMockTeamcenterPort(),
        oa: createMockOaPort(),
        tencentRtx: createMockTencentRtxPort(),
        outlook: createMockOutlookPort(),
      },
    });

    expect(result.status).toBe('completed');
    expect(result.receipts).toHaveLength(3);

    const trace = await readFile(result.tracePath, 'utf8');
    expect(trace).toContain('run.started');
    expect(trace).toContain('confirmation.received');
    expect(trace).toContain('run.completed');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- tests/runtime/runtime.test.ts`
Expected: FAIL because runtime files do not exist.

- [ ] **Step 3: Implement minimal runtime**

Implement:

- Run id creation.
- Teamcenter lookup.
- Confirmation request.
- OA draft, Tencent RTX notification, Outlook meeting execution after confirmation.
- JSONL trace events.
- Clear blocked result when confirmation returns false.

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test -- tests/runtime/runtime.test.ts`
Expected: PASS.

- [ ] **Step 5: Add redaction test and implementation**

Test should call `redactForTrace({ token: 'secret', nested: { password: 'pw' }, ok: true })` and expect token/password values to be `[REDACTED]`.

- [ ] **Step 6: Run runtime tests**

Run: `npm test -- tests/runtime`
Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add src/runtime tests/runtime
git commit -m "feat: add runtime confirmation and trace logging"
```

---

### Task 5: Pi SDK Adapter Boundary

**Files:**
- Create: `src/ports/model.ts`
- Create: `src/adapters/pi/README.md`
- Create: `src/adapters/pi/pi-model-port.ts`
- Test: `tests/adapters/pi/pi-model-port.test.ts`

**Interfaces:**
- Produces: `ModelPort.decide(state, signal): Promise<Decision>`, `createPiModelPort(options): ModelPort`.

- [ ] **Step 1: Write failing boundary test**

```ts
import { describe, expect, it } from 'vitest';
import { createPiModelPort } from '../../../src/adapters/pi/pi-model-port';

describe('Pi model adapter boundary', () => {
  it('exposes a ModelPort without leaking SDK-specific types to callers', async () => {
    const model = createPiModelPort({ mode: 'disabled' });
    const decision = await model.decide({
      runId: 'run-1',
      userRequest: 'Check material P-10086',
      step: 0,
      observations: [],
      pendingIntents: [],
      confirmedIntentIds: [],
    });

    expect(decision.type).toBe('fail');
    expect(decision.reason).toBe('pi_adapter_disabled');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- tests/adapters/pi/pi-model-port.test.ts`
Expected: FAIL because model port files do not exist.

- [ ] **Step 3: Write adapter mapping note**

Create `src/adapters/pi/README.md` with:

```md
# Pi Adapter Mapping

SDK baseline: `@earendil-works/pi-coding-agent@0.84.2`.

The core imports only `ModelPort` from `src/ports/model.ts`.

Mapping:

- `ModelPort.decide(state, signal)` maps to a future Pi `AgentSession.prompt(...)` call.
- SDK events are normalized before entering runtime state.
- Tool calls are not directly executed by Pi SDK in the MVP; runtime validates `ToolIntent` first.
- `AgentSession.abort()` maps to future run cancellation.
- Usage events will become trace events when the real session adapter is enabled.
- Malformed SDK output maps to `Decision { type: 'fail', reason: 'malformed_model_decision' }`.

The MVP ships with disabled mode so the runtime remains deterministic until credentials and provider configuration are available.
```

- [ ] **Step 4: Implement disabled Pi model port**

Create `src/ports/model.ts` and `src/adapters/pi/pi-model-port.ts` so the adapter compiles, imports the SDK package only inside adapter code, and returns a typed disabled failure decision when not configured.

- [ ] **Step 5: Run test to verify it passes**

Run: `npm test -- tests/adapters/pi/pi-model-port.test.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/ports/model.ts src/adapters/pi tests/adapters/pi
git commit -m "feat: add pi model adapter boundary"
```

---

### Task 6: CLI Golden Run

**Files:**
- Create: `src/cli/main.ts`
- Test: `tests/cli/main.test.ts`
- Modify: `package.json`

**Interfaces:**
- Consumes: `runEmployeeAssistant`.
- Produces: executable CLI via `npm run dev -- "..."`.

- [ ] **Step 1: Write failing CLI test**

```ts
import { execa } from 'execa';
import { describe, expect, it } from 'vitest';

describe('CLI golden run', () => {
  it('prints a completed summary for the sample request with auto-confirm', async () => {
    const result = await execa('npm', ['run', 'dev', '--', '--yes', 'Check material P-10086 and arrange a review']);

    expect(result.stdout).toContain('completed');
    expect(result.stdout).toContain('OA-DRAFT-P-10086');
    expect(result.stdout).toContain('RTX-MSG-001');
    expect(result.stdout).toContain('OUTLOOK-MEETING-001');
    expect(result.stdout).toContain('Trace:');
  });
});
```

- [ ] **Step 2: Add missing test dependency before running**

Modify `package.json` devDependencies:

```json
"execa": "^9.5.0"
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npm test -- tests/cli/main.test.ts`
Expected: FAIL because `src/cli/main.ts` does not exist.

- [ ] **Step 4: Implement CLI**

Implement:

- Parse `--yes`.
- Read remaining args as the user request.
- Use mock ports.
- Print confirmation plan when not using `--yes`.
- Print final status, summary, receipts, and trace path.

- [ ] **Step 5: Run CLI test**

Run: `npm test -- tests/cli/main.test.ts`
Expected: PASS.

- [ ] **Step 6: Run all verification**

Run:

```bash
npm test
npm run typecheck
npm run dev -- --yes "Check material P-10086 and arrange a review"
```

Expected:

- All tests pass.
- TypeScript typecheck passes.
- CLI prints completed flow and trace path.

- [ ] **Step 7: Commit**

```bash
git add src/cli tests/cli package.json
git commit -m "feat: add employee assistant cli"
```

---

## Self-Review

Spec coverage:

- Real business MVP scenario: covered by Tasks 2, 4, and 6.
- CLI product surface: covered by Task 6.
- Teamcenter, OA, Tencent RTX, Outlook mock adapters: covered by Task 3.
- Side-effect confirmation: covered by Task 4.
- JSONL trace: covered by Task 4.
- Pi SDK adapter isolation and mapping note: covered by Task 5.
- Core independent from SDK: covered by Tasks 1, 2, and 5.
- Failure behavior: covered partly in Task 2 and Task 4; add more cases during execution if runtime tests expose gaps.
- Acceptance criteria: covered across Tasks 1 through 6.

Placeholder scan:

- No `TBD` or open implementation placeholders.
- Adapter implementation details are deferred only for real enterprise integrations, which are explicit non-goals.

Type consistency:

- `ToolIntent`, `AgentState`, `Decision`, `Observation`, and `AgentResult` are defined in Task 1 and reused consistently.
- Port names match the spec and adapter task names.

