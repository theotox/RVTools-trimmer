# RVtools XLSX Cleaner

## Goal
Reduce bloated RVtools xlsx reports down to only infrastructure-sizing-relevant columns.

## What's Built

### `clean_rvtools.py`
- Python 3 script using `openpyxl`
- Reads RVtools `.xlsx`, outputs `<name>_cleaned.xlsx` alongside it
- Accepts multiple files as CLI args
- Silently skips missing columns (handles RVtools version differences)
- Skips sheets not in the keep-list entirely
- Supports **both** RVtools naming conventions (see below)

### Two RVtools Export Formats Supported

**Format A (your file — `v` prefix):** `vInfo`, `vCPU`, `vMemory`, `vDisk`, `vNetwork`, `vCluster`, `vHost`, `vDatastore`, `vSnapshot`, `vTools`

**Format B (legacy/other options):** `VMs`, `VMDisks`, `VMNetwork`, `VMemory`, `Hosts`, `Clusters`, `Datastores`, `VMHardware`, `VMTools`, `Snapshots`, `HostHardware`

### Sheets & Columns Kept (Format A — actual column names from your file)

| Sheet | Original | Kept | Columns |
|---|---|---|---|
| vInfo | 95 | 13 | VM, Powerstate, CPUs, Memory, Active Memory, NICs, Disks, Total disk capacity MiB, Provisioned MiB, In Use MiB, Resource pool, Folder, Primary IP Address |
| vCPU | 34 | 11 | VM, Powerstate, CPUs, Sockets, Cores p/s, Reservation, Entitlement, DRS Entitlement, Cluster, Host, OS according to the VMware Tools |
| vMemory | 38 | 5 | VM, Size MiB, Active, Cluster, Host |
| vDisk | 45 | 9 | VM, Disk, Capacity MiB, Thin, Disk Mode, Sharing mode, Controller, Cluster, Host |
| vNetwork | 32 | 8 | VM, NIC label, Adapter, Network, Connected, Mac Address, Type, IPv4 Address |
| vCluster | 35 | 10 | Name, NumHosts, TotalCpu, NumCpuCores, NumCpuThreads, TotalMemory, HA enabled, AdmissionControlEnabled, DRS enabled, VI SDK Server |
| vHost | 74 | 16 | Host, Datacenter, Cluster, CPU Model, Speed, # CPU, Cores per CPU, # Cores, # Memory, # VMs total, # VMs, # vCPUs, vCPUs per Core, vRAM, ESX Version, VMotion support |
| vDatastore | 32 | 10 | Name, Type, Capacity MiB, Provisioned MiB, In Use MiB, Free MiB, Free %, # VMs, # Hosts, Cluster name |
| vSnapshot | 26 | 6 | VM, Name, Date / time, Size MiB (total), State, Cluster |
| vTools | 34 | 6 | VM, VM Version, Tools, Tools Version, Upgradeable, Cluster |

### Dropped entirely (Format A)
vPartition, vCD, vUSB, vSource, vRP, vHBA, vNIC, vSwitch, vPort, dvSwitch, dvPort, vSC_VMK, vMultiPath, vLicense, vFileInfo, vHealth, vMetaData

### Environment
- Python 3 venv at `./venv`
- Dependencies: `openpyxl` (see `requirements.txt`)

### Usage
```bash
./venv/bin/python clean_rvtools.py rvtools-export.xlsx
```

### Bugfixes Applied
1. **Sheet naming mismatch** — RVtools uses `vInfo`, `vCPU`, etc. not `VMs`, `Hosts`. Added both naming schemes.
2. **Empty workbook crash** — `openpyxl` refuses to save a workbook with 0 sheets. Added guard.
3. **`max_column` hangs on read-only worksheet** — replaced with `iter_rows` + header-row iteration.

### Potential Future Changes
- Add/remove columns by editing `KEEP_COLUMNS` dict
- Add filtering (e.g., exclude powered-off VMs, templates)
- Add aggregation/summary sheet (totals per cluster, per datacenter)
- Add sorting (e.g., by cluster, by power state)
- Column renaming for cleaner output (e.g., "# CPU" -> "CPUs")
