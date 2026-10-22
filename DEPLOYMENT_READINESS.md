# Deployment readiness plan

This document describes the controls required before this application should
be used to make fleet-wide changes to live network devices. The current
application is a useful Windows desktop utility and should be treated as a
controlled pilot until these gates are met.

## Current baseline

The application provides a local dashboard, user roles, optional HTTPS and
client restrictions, device monitoring, configuration backups, readiness
checks, an IOS-XE install workflow, scheduling, and post-upgrade checks. CI
currently performs a static Python check, a startup smoke test, and a Windows
packaging dry run. The application checks the Cisco-published MD5 checksum during image upload,
and CI includes unit and dependency-audit jobs. Dashboard reload concurrency is configurable per site
(1–10), with queued switches and a stop-on-failure option; CLI upgrades support
configurable parallel waves and stop-on-failure by default.

These capabilities are valuable, but do not by themselves demonstrate safe
operation at fleet scale. In particular, each supported platform and IOS-XE
release combination needs lab validation, and an upgrade tool must fail safely
when a device is unreachable, low on storage, running an unexpected image, or
partially upgraded.

## Release gates

### 1. Security and access control

- Complete an independent threat model and code review of authentication,
  authorization, session expiry/invalidation, CSRF, input validation, and
  allowed-client enforcement.
- Verify password hashing parameters, secret encryption and recovery behavior,
  TLS certificate validation, and protection of backups, logs, and local
  settings on Windows.
- Add account lockout or throttling, secure session cookies, and a supported
  credential rotation/recovery process.
- Define a vulnerability disclosure and patch process.

### 2. Safe upgrade orchestration

- Add cryptographic signature verification where Cisco publishes a verifiable
  signature, and keep checking each image against Cisco's published MD5 checksum.
- Make pre-check results explicit and block upgrades on failed or stale checks.
- Add controlled canary batch sizing, maintenance windows, approval gates, and
  a fleet-wide stop control. Per-site concurrency limits and stopping queued
  switches after a failure are available, but must be load-tested at the
  target fleet size.
- Persist a per-device state machine so restart/crash recovery cannot
  accidentally repeat an unsafe step.
- Define rollback and recovery actions for every supported model/release path;
  do not treat an automatic rollback as available unless the device path has
  been demonstrated in a lab.
- Validate post-checks beyond version match: reachability, stack/route state,
  critical services, and operator-defined health checks.

### 3. Operations and scale

- Move operational state from ad-hoc local files to a transactional,
  versioned data store with backup/restore procedures and retention policies.
- Add structured, tamper-evident audit events for who approved and executed
  each action, target devices, image digest, timestamps, and results. Redact
  credentials and sensitive command output.
- Define supported device counts, concurrent jobs, response-time objectives,
  log retention, alerting, and recovery objectives; load-test those limits.
- Provide a signed, reproducible installer, software bill of materials,
  dependency vulnerability monitoring, release notes, and a rollback path for
  application updates.

### 4. Verification and release

- Build unit and integration coverage around parsing, authorization, upgrade
  state transitions, interruption/recovery, and failure paths using mocked
  devices and a representative hardware lab.
- Run security scanning and dependency review in CI; protect the release branch
  with required reviews and successful CI.
- Test every supported switch family and IOS-XE train in a lab, then run a
  documented canary and staged rollout before broad deployment.
- Maintain an operator runbook with prerequisites, pre-change evidence,
  approval record, abort criteria, escalation path, and recovery steps.

## Suggested delivery sequence

1. **Pilot hardening:** threat model, independent security review, image
   integrity verification, safer auth/session controls, and documented lab
   test matrix.
2. **Change safety:** durable job state, explicit approvals, canary batches,
   stop control, richer post-checks, and recovery drills.
3. **Operational scale:** transactional storage, centralized audit/monitoring,
   measured capacity limits, signed releases, and disaster recovery.
4. **Deployment acceptance:** pass security review, hardware matrix, failure
   injection, load targets, and staged rollout criteria for the target
   environment.

Readiness depends on the target deployment. Passing CI or completing this
checklist does not guarantee safe operation without change-management
procedures and validation on the actual device fleet.
