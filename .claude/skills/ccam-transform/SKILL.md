# CCAM Transformation
## Purpose
Transform an official CCAM Excel source file into the text format required for loading into GraphTalk.

## Input
The input is the CCAM Excel artifact referenced by the active workflow.

The source artifact must not be modified.

## Output
Create a candidate GraphTalk import file.

The generated file is a candidate only. It must not be copied to the GraphTalk loading directory or submitted to GTMonitor.

## Transformation Rules
TODO: Add the existing manual CCAM-to-GraphTalk transformation rules here.

## Output Location
Write the generated candidate file to the project's review directory.

## Restrictions
Do not modify the original CCAM source file.
Do not load anything into GraphTalk.
Do not contact GTMonitor.
Do not mark the workflow as completed.
Do not invent missing CCAM values.
If required source information is ambiguous or missing, report the problem instead of guessing.