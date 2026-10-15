# AI-Enabled CCAM Loading into GraphTalk

## Purpose

This repository implements a human-in-the-loop workflow for detecting new CCAM releases, transforming the official source data into a GraphTalk-compatible load file, obtaining human review, and triggering the existing GraphTalk CCAM batch through GTMonitor.

The design deliberately separates deterministic orchestration from AI reasoning:

- **PowerShell** owns scheduling, state, hashing, retries, reminders, file handling, and workflow transitions.
- **AI agents** interpret changed source content, perform the CCAM transformation, and analyse error logs when deterministic processing fails.
- **A human reviewer** approves the generated artifact or edits and submits the final version.
- **A connector** abstracts notifications and human approval/submission so the workflow is not coupled to Outlook or any specific email provider.
- **GTMonitor** listens for batch requests and executes the existing GraphTalk CCAM load batch.

## Problem

The current CCAM loading process requires manual work to:

1. Check the official Ameli CCAM download page for a new release.
2. Retrieve the relevant source data.
3. Review CCAM codes, labels, tariffs, and other required source values.
4. Transform those values into the text format expected by GraphTalk.
5. Review the transformed artifact.
6. Move the accepted file to the GraphTalk input location.
7. Trigger the existing CCAM load batch.
8. Investigate execution errors when loading fails.

The objective is to automate repeatable work while retaining an explicit human control point before any generated artifact is loaded.

## Source

Current CCAM information is retrieved from the configured official Ameli CCAM download page:

`https://www.ameli.fr/accueil-de-la-ccam/telechargement/index.php`

The URL must be deployment configuration rather than a value embedded throughout the implementation.

## Architecture

```text
                         DAILY SCHEDULE
                              |
                              v
                    +-------------------+
                    | PowerShell        |
                    | Workflow Controller|
                    +---------+---------+
                              |
                 Download Ameli HTML
                              |
                              v
                        SHA-256 hash
                              |
                +-------------+-------------+
                |                           |
             unchanged                    changed
                |                           |
                |                           v
                |                 Low-cost AI interpreter
                |                           |
                |                   current CCAM version
                |                   + source artifacts
                |                           |
                |                 +---------+---------+
                |                 |                   |
                |             already known          new
                |                 |                   |
                |                 |                   v
                |                 |          supersede pending
                |                 |          older release
                |                 |                   |
                |                 |                   v
                |                 |          download source files
                |                 |                   |
                |                 |                   v
                |                 |          Transformation Agent
                |                 |                   |
                |                 |                   v
                |                 |          Candidate Artifact
                |                 |                   |
                |                 |                   v
                |                 |          Notification Connector
                |                 |                   |
                |                 |                   v
                |                 |             Human Review
                |                 |             |          |
                |                 |           accept      edit
                |                 |             |          |
                |                 |             +----+-----+
                |                 |                  |
                |                 |                  v
                |                 |          Final Artifact Submitted
                |                 |                  |
                |                 |                  v
                |                 |          Stage Approved Artifact
                |                 |                  |
                |                 |                  v
                |                 |          Notify GTMonitor
                |                 |                  |
                |                 |                  v
                |                 |          GTMonitor executes batch
                |                 |                  |
                |                 |                  v
                |                 |          Inspect execution log
                |                 |          /        |         \
                |                 |      missing    error    completed
                |                 |         |         |          |
                |                 |         v         v          v
                |                 |      technical   AI error   COMPLETED
                |                 |      intervention analysis
                |                 |                   |
                +-----------------+-------------------+
                                  |
                             persisted state
```

## Core Design Rules

### Deterministic workflow control

AI agents do not own workflow state transitions. PowerShell evaluates deterministic conditions and persists the resulting state.

Examples include:

- comparing page hashes;
- checking dates;
- determining whether a reminder is due;
- determining whether an artifact has been superseded;
- checking for an expected GTMonitor log;
- detecting the configured GTMonitor completion marker;
- updating workflow state only after the relevant condition has been verified.

### AI is used for reasoning-heavy work

AI is used for:

1. interpreting the Ameli page after its content changes;
2. identifying the current CCAM version and required source artifacts;
3. transforming CCAM source data according to the project prompt/instructions;
4. analysing GTMonitor/GraphTalk error logs when a load fails.

The CCAM transformation rules are maintained in the transformation agent prompt/project instructions rather than duplicated in workflow orchestration code.

### Mandatory human review

The transformation agent produces a **candidate artifact**, never a directly loadable approved artifact.

The reviewer can:

- approve the candidate unchanged; or
- edit the candidate and submit the corrected artifact.

In both cases, the artifact submitted through the connector becomes the **final artifact**.

The loader must operate on that exact submitted artifact.

### Connector abstraction

The core workflow does not depend on Outlook or another specific communication product.

The configured connector is responsible for:

- issuing the initial review notification;
- issuing review reminders;
- allowing the reviewer to approve or submit a modified final artifact;
- notifying the configured technical contact when intervention is required;
- optionally notifying stakeholders of load success/failure.

The workflow consumes connector-level events, such as:

```text
FINAL_ARTIFACT_SUBMITTED
```

rather than provider-specific concepts such as an email reply.

## Workflow States

The initial state machine is:

```text
DETECTED
   |
   v
TRANSFORMING
   |
   v
AWAITING_REVIEW
   |
   v
APPROVED
   |
   v
LOADING
  / | \
 /  |  \
FAILED | COMPLETED
       |
MANUAL_INTERVENTION_REQUIRED
```

Any non-completed version may transition to:

```text
SUPERSEDED
```

### State definitions

| State | Meaning |
|---|---|
| `DETECTED` | A previously unknown CCAM release has been identified. |
| `TRANSFORMING` | Source artifacts are being transformed. |
| `AWAITING_REVIEW` | A candidate exists and the human reviewer has been notified. |
| `APPROVED` | The reviewer has submitted the final artifact. |
| `LOADING` | The approved artifact has been staged and GTMonitor execution has been requested. |
| `COMPLETED` | The expected GTMonitor execution log confirms successful completion. |
| `FAILED` | Transformation or loading failed and diagnostic information is available. |
| `MANUAL_INTERVENTION_REQUIRED` | Expected deterministic evidence, such as the GTMonitor execution log, is missing or cannot be evaluated safely. |
| `SUPERSEDED` | A newer CCAM release replaced this release before completion. |

## Supersession Rule

A newer CCAM version completely replaces an older version that has not completed successfully.

Example:

```text
V84.00 = COMPLETED
V85.00 = AWAITING_REVIEW
V86.00 = detected
```

The resulting workflow is:

```text
V84.00 = COMPLETED
V85.00 = SUPERSEDED
V86.00 = DETECTED / active workflow
```

After V85.00 becomes `SUPERSEDED`:

- review reminders stop;
- approval of V85.00 must not initiate loading;
- V85.00 artifacts remain available for traceability;
- processing continues with V86.00.

Supersession is enforced by deterministic workflow logic, not by an AI decision.

## Daily Detection Workflow

A Windows scheduled task executes the PowerShell controller once per day.

The controller:

1. Downloads the configured Ameli page using `curl.exe`.
2. Computes SHA-256 for the downloaded HTML.
3. Loads the persisted state.
4. Processes any due human-review reminder.
5. Compares the new page hash with the stored page hash.
6. If unchanged, terminates without invoking the source-interpreter AI.
7. If changed, invokes a low-cost AI model to interpret the current page.
8. Requires a structured response identifying the current CCAM version and required source artifacts.
9. Compares the detected version with workflow state.
10. Starts a new workflow only when a genuinely new release is identified.

A page-hash change means only **the source page changed**. It does not by itself mean a new CCAM version was released.

## Source Interpreter Contract

The AI source interpreter must return machine-readable structured output.

Illustrative response:

```json
{
  "version": "V85.00",
  "sourceArtifacts": [
    {
      "type": "CCAM_SOURCE",
      "url": "https://example/source"
    }
  ],
  "confidence": "high"
}
```

The controller validates required fields before creating or changing workflow state.

The interpreter must identify the **current** CCAM release rather than historical releases also present on the webpage.

## Transformation

For a newly detected release:

1. The required official source artifacts are downloaded.
2. State becomes `TRANSFORMING`.
3. The transformation agent receives the source artifacts and project transformation instructions.
4. The agent generates the GraphTalk-compatible candidate file.
5. The candidate artifact is hashed and placed in the review area.
6. State becomes `AWAITING_REVIEW`.
7. The connector sends the review request.

Detailed transformation rules belong in the transformation agent prompt/project instructions.

## Human Review and Reminders

The reviewer receives the candidate through the configured connector.

The reviewer either approves it unchanged or edits and submits a replacement final artifact.

The controller records the final artifact and its SHA-256 before progressing to loading.

### Reminder policy

When review is requested, state records a date such as:

```json
{
  "reviewRequestedAt": "2026-10-15T08:00:00Z",
  "nextReminderAt": "2026-10-20T08:00:00Z",
  "reminderCount": 0
}
```

During the daily run, when:

```text
status == AWAITING_REVIEW
AND current time >= nextReminderAt
```

the connector sends a reminder, `reminderCount` is incremented, and the next configured reminder date is calculated.

A superseded workflow never sends additional review reminders.

The reminder interval and escalation policy are configuration values.

## GraphTalk Loading Through GTMonitor

After the final artifact is submitted:

1. State becomes `APPROVED`.
2. The controller confirms that the version remains active and has not been superseded.
3. The exact approved artifact is copied to the predefined GraphTalk input directory.
4. A batch request is sent to GTMonitor.
5. GTMonitor listens for the request and executes the existing CCAM batch.
6. State becomes `LOADING` while execution outcome is pending.

## Determining Load Outcome

The system uses GTMonitor/GraphTalk execution logs as the authoritative evidence of the batch result.

### Successful load

The expected execution log exists and contains the configured deterministic completion marker.

The controller then:

1. sets the workflow to `COMPLETED`;
2. records the successfully loaded version as `processedVersion`.

### Missing execution log

If the expected log is not available according to the configured checking/retry policy:

1. state becomes `MANUAL_INTERVENTION_REQUIRED`;
2. the connector notifies the configured technical contact;
3. `processedVersion` is not changed.

### Execution error

When error-specific logs are available:

1. state becomes `FAILED`;
2. relevant error logs are collected;
3. an AI diagnostic agent analyses those logs;
4. the connector sends the technical contact the failure information and AI-assisted analysis;
5. `processedVersion` is not changed.

The AI diagnostic agent does **not** decide whether the load succeeded. Success is determined using the configured deterministic GTMonitor log marker.

## Persisted State

Initial implementation uses JSON because the workflow has multiple related values and the state must be usable by PowerShell and readable by developers.

Illustrative state:

```json
{
  "processedVersion": "V84.00",
  "detectedVersion": "V85.00",
  "status": "AWAITING_REVIEW",
  "pageHash": "<sha256>",
  "detectedAt": "2026-10-15T08:00:00Z",
  "reviewRequestedAt": "2026-10-15T08:10:00Z",
  "nextReminderAt": "2026-10-20T08:10:00Z",
  "reminderCount": 0,
  "candidateArtifact": "review/CCAM_V85_candidate.txt",
  "candidateArtifactHash": "<sha256>",
  "finalArtifact": null,
  "finalArtifactHash": null,
  "lastError": null
}
```

`processedVersion` changes only after successful load completion has been verified.

## Initial Repository Structure

```text
ccam-automation/
|
|-- README.md
|-- CLAUDE.md
|
|-- detector/
|   `-- check-ccam.ps1
|
|-- state/
|   |-- workflow.json
|   `-- archive/
|
|-- config/
|
|-- input/
|
|-- work/
|
|-- review/
|
|-- approved/
|
|-- logs/
|
`-- .claude/
    `-- skills/
```

## Configuration

Deployment-specific values must not be hard-coded into agent prompts or developer-specific paths.

Initial configuration should include:

```text
CCAM source URL
review reminder interval
reviewer connector/destination
technical contact connector/destination
GraphTalk input directory
GTMonitor batch request configuration
GTMonitor execution-log location/pattern
GTMonitor success marker
working directories
AI model used for source interpretation
AI model used for transformation
AI model used for error-log analysis
```

## Idempotency and Safety Requirements

The workflow must be safe to run repeatedly.

At minimum:

- An unchanged page does not invoke AI unnecessarily.
- The same CCAM release does not start multiple transformations.
- The initial review request is not resent every day.
- Reminders are controlled using persisted dates.
- A superseded version cannot be loaded.
- A completed version is not automatically reloaded.
- `processedVersion` changes only after verified load success.
- Missing logs cannot be interpreted as success.
- AI analysis cannot override deterministic load-status checks.
- The exact human-submitted artifact is the artifact staged for GraphTalk.
- Workflow state must survive machine/process restarts.
- Secrets and authentication credentials must not be stored in repository files or logs.

## Definition of Done for a CCAM Release

A release is considered `COMPLETED` only after all of the following have occurred:

1. A new release was identified from the configured CCAM source.
2. Required source artifacts were retrieved.
3. The transformation produced a candidate artifact.
4. A human reviewer submitted the final artifact.
5. The exact final artifact was staged for loading.
6. GTMonitor received the CCAM batch request.
7. GTMonitor executed the batch.
8. The expected execution log was located.
9. The configured completion marker confirmed successful execution.
10. Workflow state was updated and `processedVersion` was advanced.

Anything before step 9 is an in-progress or failed workflow, not a successfully loaded CCAM release.

## Remaining Implementation Decisions

The architecture is sufficiently defined to begin implementation. Remaining decisions are implementation/configuration concerns rather than fundamental workflow questions:

- connector technology used in the first prototype;
- exact connector event/payload for final artifact submission;
- exact GTMonitor batch-request interface;
- exact execution-log naming/location and completion marker;
- retry/check policy while waiting for the GTMonitor execution log;
- escalation rules after repeated unanswered review reminders;
- production persistence mechanism if JSON becomes insufficient for the deployment environment.
