# Service Voucher Management System

A Windows desktop application for managing service vouchers, recipients, users,
technician work, commissions, PDF vouchers, backups, and CSV exports.

## Current capabilities

- Secure first-run administrator setup with no hardcoded default password.
- Bcrypt password hashing, forced temporary-password changes, login throttling,
  active-account checks, and role-based permissions.
- Create, search, edit, update, refresh, and delete service vouchers with
  database-backed pagination (100 records per page).
- Edit individual numeric voucher IDs and sort them numerically in either direction.
- Export every voucher matching the current filters to one CSV, across all pages.
- Atomic, monotonic voucher allocation with a configurable minimum base number;
  deleted or renamed high IDs are not silently reused.
- Staff and recipient management.
- User administration with administrator safeguards.
- Commission records linked to staff and optionally to voucher IDs. Voucher billing,
  technician, amount, and commission changes automatically synchronize linked records.
- Filtered commission pagination, numeric ID sorting, filtered totals, and full CSV
  export across every matching page.
- Complete voucher PDFs containing customer, device, service, technician, billing,
  and signature details.
- SQLite schema migrations for databases created by older versions.
- Audit logging for security-sensitive and business-data changes, with pagination.
- ZIP backups containing a consistent database snapshot, PDFs, and staff attachments.
- Validated, backward-compatible restore with an automatic pre-restore backup.
- Windows PyInstaller build through GitHub Actions.

The feature-rich November 2025 implementation is preserved on the
`archive/feature-complete-2025-11-26` branch. The active version uses maintainable,
separate modules rather than the former 5,000-line `main.py`.

## Requirements

- Windows 10 or Windows 11 for normal desktop use.
- Python 3.11 or later when running from source.

## Run from source

```powershell
py -3.11 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

On first launch, the application asks you to create the first administrator.
There are no default credentials.

## Roles

| Role | Main permissions |
|---|---|
| Administrator | All voucher, user, staff, commission, backup, audit, and settings functions |
| Sales assistant | Create/edit vouchers, update status, commissions, and CSV exports |
| Technician | Update voucher status/solution and export filtered records |
| User | View vouchers, open PDFs, and export filtered records |

Administrators can customize an account's role and active status from **Users**.
New accounts receive a temporary password and must change it at next login.

## Voucher and commission synchronization

Saving a voucher with a reference bill and a technician that matches a Staff record
by staff ID or name creates or updates its linked commission automatically. `INV...`
reference numbers are classified as invoices; other references use cash-bill type.
Changing a voucher ID also updates commission links in the same transaction. If the
required link fields are later cleared, voucher-managed commission rows are removed,
while manually created commission rows and their links are preserved. Deleting a
voucher removes its generated commission and safely unlinks any manual commission.

## Operational data

Source builds store data beside the source files. Packaged Windows builds store
writable data in `%LOCALAPPDATA%\TONY.COM\ServiceVoucherApp` so installed or MSIX
application directories remain read-only:

- `vouchers.db` — SQLite database
- `pdfs/` — generated voucher PDFs
- `staffs/` — staff attachments
- `logs/app.log` — rotating application log
- `backups/` — automatic pre-restore backups

To store data elsewhere, set `SVMS_DATA_DIR` before launching:

```powershell
$env:SVMS_DATA_DIR = "D:\SVMS-Data"
python main.py
```

The following optional environment variables customize business details:

- `SVMS_SHOP_NAME`
- `SVMS_SHOP_ADDR`
- `SVMS_SHOP_TEL`
- `SVMS_DB_FILE`

Back up the application through **Backup** before moving, updating, or replacing a
production installation.

## Backup and restore

**Backup** creates a ZIP containing a transactionally consistent database snapshot
plus PDFs, staff attachments, and any legacy `images/` attachments that remain in use.

**Restore** validates archive paths, checks SQLite integrity, and creates a
`pre_restore_*.zip` copy of current data before restoring. The application closes
after a successful restore so the restored account and permission state is applied
on the next launch. Current versioned backups, earlier format-1 backups, and legacy
metadata-free ZIPs with `vouchers.db` at the root or inside one enclosing application
folder are accepted. Restored databases are migrated in place without renumbering
existing vouchers.

Keep backups on a protected drive because they contain customer and staff data.

## Tests

```powershell
pip install -r requirements-dev.txt
python -m pytest
python -m compileall -q .
```

GitHub Actions runs the tests before producing the Windows build artifact.

## Windows build

Push to a branch or manually run the **SVMS Tests and Windows Build** workflow.
After a successful run, download the `ServiceVoucherApp-onedir` artifact, extract
the ZIP, and run `ServiceVoucherApp.exe`.

## Repository safety

The `.gitignore` excludes the database, PDFs, logs, backups, staff attachments, and
build outputs. Do not force-add those files: they may contain personal or business
information.

See [SECURITY.md](SECURITY.md) for deployment and vulnerability guidance.
