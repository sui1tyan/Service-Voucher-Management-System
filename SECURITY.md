# Security guidance

## Supported version

Security fixes are applied to the current `main` branch. The archived legacy branch
is retained only for recovery and feature reference.

## Reporting a vulnerability

Do not open a public issue containing credentials, customer details, database files,
or exploitable security information. Contact the repository owner privately or use
a private GitHub security advisory.

## Deployment checklist

- Create a unique first administrator password.
- Use separate named accounts instead of sharing an administrator login.
- Give each account only the role it needs.
- Store `SVMS_DATA_DIR` on a drive protected by Windows account permissions.
- Create regular backups and protect them like the live customer database.
- Never commit `vouchers.db`, PDFs, logs, staff attachments, or backup ZIP files.
- Review the audit log and deactivate accounts that are no longer required.
- Install only build artifacts produced by a successful repository workflow.
