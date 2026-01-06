# Assumptions

Technical assumptions, constraints, and dependencies for the Protection
Assessment script.

## PowerFactory Model Requirements

### Study Case Configuration

- An active study case must be selected before script execution
- The study case must contain at least one active external grid (ElmXnet)
- Feeders must be radial (mesh feeders are automatically excluded)
- Lines placed out of service may cause feeders to be excluded from
  the selection list if they serve as open points

### Protection Device Configuration

#### Relays (ElmRelay)

- Must have a relay type (`typ_id`) assigned
- Must have a CT (`cpCt`) connected
- CT phase count must match the relay measurement element phase count:
  - 3-phase CT requires 3-phase measurement type
  - 1-phase CT requires 1-phase measurement type
- Relay must be located in a cubicle (`StaCubic`)
- Relay must not be out of service

#### Fuses (RelFuse)

- Must have a fuse type with valid time-current characteristic curve
- Fuse type must use Hermite Polynomial curve representation (type 6)
- Must be connected to an energized terminal
- Distribution transformer fuses are excluded from protection section
  analysis (only line fuses are considered as protection devices)

#### Auto-Reclosers

- Auto-reclose elements (`RelRecl`) define the trip sequence
- The `oplockout` attribute determines total trips before lockout
- Block logic in the recloser type controls which elements are active
  for each trip in the sequence
- Relays without auto-reclose are treated as single-trip devices

### Network Topology

- Terminals must have nominal voltage > 1 kV to be included in
  fault study results
- Line construction type is determined by:
  - `TypGeo`: Tower geometry (overhead)
  - `TypLne`: Standard line type
  - `TypCabsys`: Cable system (underground)
  - Line type name containing "SWER": Single Wire Earth Return
- Floating terminals (end-of-line with single connection) are detected
  and fault levels calculated via line fault location

## Regional Differences

### SEQ Models

| Aspect | SEQ Assumption |
|--------|----------------|
| Load representation | Aggregated loads (`ElmLod`) at substations |
| Fault resistance | 0 Ω for all earth faults |
| Primary reach threshold | 2.0 |
| Backup reach threshold | 1.3 |
| Switch identification | `StaSwitch` objects |
| Fuse selection | Energex fuse mapping tables |

### Regional Models

| Aspect | Regional Assumption |
|--------|---------------------|
| Load representation | Individual transformers (`ElmTr2`) |
| Fault resistance (OH) | 50 Ω for overhead lines |
| Fault resistance (UG) | 10 Ω for underground cables |
| Primary reach threshold | 1.7 |
| Backup reach threshold | 1.3 |
| Switch identification | `ElmCoup` and `StaSwitch` with specific prefixes |
| Fuse selection | Ergon Energy fuse mapping tables |

### Region Detection

The script determines the region by examining the derived base project
path:

- Path containing "Regional Models" → Regional
- Path containing "SEQ" → SEQ

If neither pattern matches, the script raises a `RuntimeError`.

## Fault Study Parameters

### Short-Circuit Calculation Method

- Calculation method: Complete (IEC 60909 method 3)
- Loads ignored in fault calculations
- Line capacitance ignored
- Shunts ignored
- Protection tripping current: Transient mode

### Voltage Factors

| Study Type | c-Factor |
|------------|----------|
| Maximum fault level | 1.1 |
| Minimum fault level | 1.0 |

### Fault Types Calculated

| Fault Type | Attribute Prefix | Usage |
|------------|------------------|-------|
| 3-Phase | `_3ph` | Maximum phase fault, breaking capacity |
| 2-Phase | `_2ph` | Minimum phase fault, reach factors |
| Phase-Ground | `_pg` | Earth fault protection coordination |
| Phase-Ground 10Ω | `_pg10` | Regional UG minimum (stored only) |
| Phase-Ground 50Ω | `_pg50` | Regional OH minimum (stored only) |

### System Normal Minimum

If the user-entered system normal minimum grid parameters differ from
the standard minimum parameters, a separate set of fault studies is
executed. Otherwise, minimum values are copied to system normal fields.

When system normal minimum grid impedance is zero or negative, that
grid is placed out of service for the system normal study.

## Protection Reach Calculations

### Reach Factor Definition

```
Reach Factor = Minimum Fault Current at Location / Device Pickup Setting
```

### Pickup Determination

- For relays: Highest pickup among active IDMT or instantaneous elements
  of each protection function (phase OC, earth fault, NPS)
- For fuses: Rated current × 2

### SWER Transformation

When a protection device at a higher voltage level protects a SWER
section at a lower voltage:

```
Device Current = (SWER Voltage × SWER Fault Current) / (Device Voltage × √3)
```

The device sees the fault as a 2-phase equivalent rather than earth fault.

### Active Elements by Fault Type

| Fault Type | Active Protection Functions |
|------------|----------------------------|
| 3-Phase | Phase OC only |
| 2-Phase | Phase OC + NPS |
| Phase-Ground | Phase OC + Earth Fault + NPS |

## Conductor Damage Assessment

### Energy Calculation

Total let-through energy is accumulated across all trips in the
auto-reclose sequence:

```
Total Energy = Σ (Fault Current² × Clearing Time) for each trip
```

### Clearing Time Components

- Relay operate time (from time-current characteristic)
- Circuit breaker opening time: 80 ms (fixed assumption)

### Thermal Rating Source

- Overhead lines: `Ithr` attribute from conductor type (1-second rating)
- Cable systems: Not assessed ("NA" returned)

### Pass/Fail Criteria

```
Allowable Energy = Thermal Rating² × 1 second
Result = PASS if Total Energy ≤ Allowable Energy else FAIL
```

SWER lines return "SWER" status for phase faults (not applicable).

## Fuse Selection

### Mapping Tables

Fuse selection uses lookup tables based on:

- Region (SEQ or Regional)
- Transformer kVA rating
- Line construction type (OH, UG, SWER)
- Voltage level
- Number of phases
- RMU insulation type (air or oil) - SEQ only
- Transformer impedance class (high or low) - SEQ RMU only

### Source Documents

- SEQ: Energex fuse tables
- Regional: Ergon Energy Technical Instruction TSD0019k

## Output Assumptions

### Excel File Location

Files are saved to the user's local directory, attempting paths in order:

1. `//client/c$/LocalData/{username}/` (Citrix client mapping)
2. `c:/LocalData/{username}/` (Local installation)

### PowerFactory Colour Scheme Encoding

DPL attributes (`dpl1`, `dpl2`, `dpl3`) encode pairs of assessment
results using a 4×4 grid:

```
Value = (Row Condition - 1) × 4 + Column Condition
```

Where conditions are: Pass=1, Fail=2, No Data=3, SWER=4

| DPL Attribute | Row Assessment | Column Assessment |
|---------------|----------------|-------------------|
| `dpl1` | Phase Primary Reach | Phase Backup Reach |
| `dpl2` | Earth Primary Reach | Earth Backup Reach |
| `dpl3` | Phase Conductor Damage | Earth Conductor Damage |

### Time-Current Plot Settings

- Plot scale: Logarithmic (both axes)
- Y-axis range: 0.01 to 10 seconds
- X-axis range: Auto-scaled to fault level range
- Earth fault curves: Dashed line style
- Page format: A4 Landscape

## Dependencies

### Python Packages

| Package | Purpose |
|---------|---------|
| `pandas` | DataFrame operations, Excel export |
| `openpyxl` | Excel workbook formatting |
| `tkinter` | GUI dialogs |
| `math` | Mathematical operations |
| `logging` | Error and debug logging |

### PowerFactory Version

Script developed and tested with PowerFactory 2023 and later.
Earlier versions may have API differences.

### Network Path Dependencies

- PowerFactory typing stubs: `\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping`
- Ergon global library: Required for fuse type lookups

## Known Limitations

1. **Mesh feeders excluded** - Only radial feeders can be studied
2. **Single protection section** - Devices define sections by downstream
   topology; parallel paths may cause unexpected section boundaries
3. **Fixed breaker time** - 80 ms assumed for all circuit breakers
4. **Cable thermal rating** - Not assessed (requires different methodology)
5. **Voltage regulator exclusion** - Transformers in "Regulators" folder
   are excluded from load calculations
6. **Single fuse curve type** - Only Hermite Polynomial (type 6) supported