# Protection Assessment

PowerFactory protection assessment and fault study automation tool for
distribution network analysis.

## Repository

Code under development on Gitea:
http://smartweb.myergon.local:3000/dp072/protection-assessment

## Overview

This tool performs comprehensive protection assessment studies for
distribution feeders in DIgSILENT PowerFactory, including:

- **Fault Level Studies** - Maximum and minimum fault current calculations
  at all network terminals
- **Protection Reach Analysis** - Primary and backup device reach factor
  calculations against regional thresholds
- **Conductor Damage Assessment** - Thermal withstand evaluation for
  overhead conductors under fault conditions
- **Protection Coordination Plots** - Time-overcurrent characteristic
  curves for relay coordination review

Results are exported to Excel workbooks and PowerFactory graphical outputs.

## Features

### Study Types

| Study | Description |
|-------|-------------|
| Find PowerFactory Project | Locate the master project for a given substation acronym |
| Find Feeder Open Points | Identify normally-open points on radial feeders |
| Fault Level Study (Legacy) | SEQ-specific fault study with legacy output format |
| Fault Level Study (No Relays) | Fault calculations at switch locations |
| Fault Level Study (All Relays) | Full protection assessment with configured relays |

### Optional Assessments

When running the full relay-configured study:

- **Conductor Damage Assessment** - Evaluates I²t let-through energy against
  conductor thermal ratings across the auto-reclose sequence
- **Protection Coordination Plots** - Generates time-overcurrent plots with
  upstream/downstream device curves and fault level markers

### Regional Support

The tool supports both network regions with appropriate fault impedance
assumptions:

- **SEQ Models** - Aggregated load representation, 0Ω fault resistance
- **Regional Models** - Individual transformer representation,
  construction-dependent fault resistance (50Ω OH, 10Ω UG)

## Project Structure

```
protection-assessment/
├── start.py                 # Main entry point
├── pf_config.py             # Path configuration and PowerFactory typing
├── pf_protection_helper.py  # Context managers and utilities
├── model_checks.py          # Pre-study model validation
│
├── domain/                  # Core domain models (dataclasses)
│   ├── enums.py             # Element types, fault types, construction types
│   ├── fault_data.py        # Immutable fault current containers
│   ├── feeder.py            # Feeder model
│   ├── device.py            # Protection device model
│   ├── termination.py       # Network terminal model
│   ├── line.py              # Distribution line model
│   ├── transformer.py       # Transformer/load model
│   └── utils.py             # Domain utilities
│
├── fault_study/             # Fault analysis modules
│   ├── fault_level_study.py # Main fault study orchestration
│   ├── analysis.py          # Short-circuit execution and results
│   ├── study_templates.py   # PowerFactory study configurations
│   ├── fault_impedance.py   # Construction-based impedance handling
│   └── floating_terminals.py# End-of-feeder terminal handling
│
├── relays/                  # Protection device analysis
│   ├── elements.py          # Relay/fuse retrieval and filtering
│   ├── reach_factors.py     # Reach factor calculations
│   ├── reclose.py           # Auto-reclose sequence management
│   └── current_conversion.py# Fault current to measurement conversion
│
├── devices/                 # Device type handling
│   ├── fuses.py             # Fuse creation and selection
│   └── fuse_mapping.py      # Regional fuse sizing tables
│
├── cond_damage/             # Conductor damage assessment
│   └── conductor_damage.py  # Energy calculations and damage evaluation
│
├── colour_maps/             # PowerFactory visualization
│   └── colour_maps.py       # Conditional diagram colouring
│
├── save_results/            # Output generation
│   ├── save_result.py       # Excel workbook creation
│   └── cond_dmg_results.py  # Conductor damage DataFrame formatting
│
├── oc_plots/                # Overcurrent plot generation
│   ├── plot_relays.py       # Time-overcurrent plot creation
│   └── get_rmu_fuses.py     # RMU transformer fuse specification GUI
│
├── user_inputs/             # User interface modules
│   ├── study_selection.py   # Study type selection dialog
│   └── get_inputs.py        # Feeder/device selection and grid data
│
├── fdr_open_points/         # Feeder topology analysis
│   ├── get_open_points.py   # Open point detection
│   └── fdr_open_user_input.py# Feeder selection interface
│
├── find_substation/         # Project location utility
│   └── find_sub.py          # Substation-to-project mapping
│
├── legacy_script/           # Legacy SEQ output format
│   ├── script_bridge.py     # Data structure conversion
│   └── save_results.py      # Legacy Excel formatting
│
└── config_logging/          # Logging configuration
    └── configure_logging.py # Path resolution and log setup
```

## Requirements

- DIgSILENT PowerFactory (with Python scripting enabled)
- Python packages: `pandas`, `openpyxl`, `tkinter`
- Network access to PowerFactory typing stubs (for development)

## Usage

1. Map `start.py` to a PowerFactory script object as the executable
2. Activate the desired study case in PowerFactory
3. Execute the script
4. Select the study type and feeders via the GUI dialogs
5. Results are saved to:
   - Excel workbook in the user's local directory
   - PowerFactory graphics board (for coordination plots)
   - User-defined colour schemes (for reach/damage visualization)

## Configuration

### Protection Device Requirements

All protection devices must be correctly configured in the PowerFactory
model:

- **Relays (ElmRelay)** - Must have relay type assigned with CT connected
- **Fuses (RelFuse)** - Must have fuse type with valid time-current curve
- **Measurement elements** - Phase count must match CT configuration

### External Grid Data

Users are prompted to enter or confirm external grid fault level data:

- Maximum fault level (3-phase, R/X, sequence impedance ratios)
- Minimum fault level parameters
- System normal minimum parameters (if different from minimum)

## Output Files

### Excel Workbook Contents

| Sheet | Description |
|-------|-------------|
| General Information | Study parameters, grid data, calculation settings |
| Summary Results | Device fault levels, downstream capacity, backup devices |
| {Feeder} Detailed Results | Terminal-by-terminal fault levels and reach factors |
| {Feeder} Cond Dmg Res | Conductor damage pass/fail results (if selected) |

### PowerFactory Outputs

- **Graphics Board** - Time-overcurrent coordination plots per device
- **Colour Schemes** - User-defined conditional formatting for:
  - Phase/Earth primary reach (Pass/Fail/No Data)
  - Phase/Earth backup reach (Pass/Fail/No Data)
  - Phase/Earth conductor damage (Pass/Fail/No Data/SWER)

## Maintainers

- **Primary:** dan.park@energyq.com.au