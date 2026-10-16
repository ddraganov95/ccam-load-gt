# CCAM Load Automation

## Project Purpose

This project automates the detection, transformation, human review, and loading of new CCAM releases into GraphTalk.

The current CCAM loading process is an ETL-style process in which extraction and transformation require manual work. This project automates repeatable parts of that workflow while keeping a mandatory human review step before a generated artifact can be loaded into GraphTalk.

## Architecture

The workflow is:

1. A PowerShell controller runs periodically.
2. The controller downloads the official CCAM download page.
3. The downloaded HTML is hashed and compared with the previously observed page hash.
4. If the page has not changed, no AI source analysis is performed.
5. If the page has changed, AI analyses the page and identifies:
   - the complete current CCAM version;
   - the Excel source artifact for that release.
6. The controller downloads the official Excel artifact and records its SHA-256 hash.
7. The CCAM transformation agent transforms the Excel source into a candidate GraphTalk load file.
8. A human reviewer reviews the candidate.
9. The reviewer may:
   - approve the candidate unchanged; or
   - modify it and submit a corrected final artifact.
10. The submitted artifact becomes the final approved artifact.
11. The approved artifact is copied to the configured GraphTalk input directory.
12. GTMonitor is notified that the CCAM batch should execute.
13. GTMonitor executes the existing GraphTalk CCAM load batch.
14. Execution logs are inspected to determine the result.
15. The workflow becomes COMPLETED only after a successful execution is confirmed.

## Responsibilities

### PowerShell

PowerShell is the deterministic workflow controller.

PowerShell owns:

- workflow state;
- configuration loading;
- source-page downloading;
- SHA-256 calculation;
- deciding whether source content changed;
- invoking AI tasks;
- validating structured AI responses;
- downloading source artifacts;
- reminder scheduling;
- supersession logic;
- file movement;
- GTMonitor invocation;
- deterministic inspection of execution status;
- workflow state transitions.

Do not move deterministic workflow decisions into AI prompts when they can be expressed reliably in code.

### AI

AI is used for tasks requiring interpretation or transformation.

Current AI responsibilities are:

1. Interpret changed CCAM webpage content and identify the current CCAM release and Excel artifact.
2. Transform the official CCAM Excel artifact into the candidate GraphTalk input artifact.
3. Analyse GraphTalk/GTMonitor error logs when load 