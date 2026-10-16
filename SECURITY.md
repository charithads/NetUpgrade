# Security policy

## Scope

This project can store network device credentials and run privileged IOS-XE
commands that may interrupt live network traffic. Treat the application host,
its local data, and its network path to switches as part of the network
management security boundary.

## Safe deployment baseline

- Run the manager on a dedicated, patched Windows host with a named operator
  account and full-disk encryption.
- Keep the dashboard bound to loopback unless remote access is required. If
  remote access is enabled, require HTTPS, restrict client addresses, and place
  it behind network controls that limit access to approved administrators.
- The direct manager cannot be exposed remotely. Remote binding is accepted
  only through the desktop launcher with configured users, HTTPS, and the
  explicit remote-access setting.
- Use individual, least-privilege switch accounts where the platform supports
  them. Do not reuse personal credentials or share credentials across teams.
- Limit device management access to the required management VLAN and hosts.
- Keep configuration backups and audit records access-controlled and backed
  up. They may contain sensitive topology and configuration data.
- When uploading an IOS-XE `.bin` image, paste the 128-character SHA-512
  checksum shown on Cisco's software download page. NetUpgrade checks the image
  before saving it and blocks a missing, malformed, or mismatched checksum.
  Keep a tested recovery image and out-of-band access available.
- Upgrade a lab device and a small canary group first. Require a human change
  approval, a maintenance window, pre-check evidence, post-check evidence, and
  an explicit stop decision before expanding a rollout.
- Do not expose the dashboard directly to the public internet.

## Known security limitations

This application has not undergone an independent security assessment. It is a
single-host desktop application, not a hardened multi-tenant service. Before
broad deployment, review session handling, CSRF protections, rate limiting,
secret storage and recovery, audit-log integrity, input validation, TLS
configuration, and privilege boundaries. Run it only in a controlled pilot
until those controls have been verified for the intended deployment.

## Reporting a vulnerability

Do not open a public issue containing credentials, configuration files, or
exploit details. Contact the repository owner privately to coordinate a fix.
Include the affected version, impact, and a minimal reproduction without real
device secrets.
