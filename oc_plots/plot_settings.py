import math

from pf_config import pft
from typing import Dict, List
import domain as dd

# =============================================================================
# PLOT Level 2 SETTINGS
# =============================================================================

def setup_drawing_format(setgrfpage):
    """
    Create A4 landscape page format in project settings.

    Args:
        app: PowerFactory application instance.

    Returns:
        SetFormat object configured for A4 landscape (297x210mm).
    """

    # Drawing Size
    setgrfpage.SetAttribute("iDrwFrm", 1)                           # Landscape
    # Format
    setgrfpage.SetAttribute("aDrwFrm", "A4")                        # A4


# =============================================================================
# PLOT Level 3 SETTINGS
# =============================================================================


def setup_toc_plot(pltovercurrent, f_type: str) -> None:


    # Curves
    # =============================================================================

    # Plot features
    # Show coordination results
    pltovercurrent.SetAttribute("showCoordinationResults", 0)           # No
    # Select sub curves
    pltovercurrent.SetAttribute("enableSubCurves", 0)                   # No

    # Elements


    # Drawing options
    # =============================================================================

    # Voltage reference axes
    # Current unit
    pltovercurrent.SetAttribute("currentUnit", 0)                       # Primary ampere
    # Show voltages
    pltovercurrent.SetAttribute("shownVoltageRefAxes", 2)               # User-defined only
    # Voltage
    pltovercurrent.SetAttribute("userDefVoltageRefAxis", 11)            # 11kV

    # Show relays
    if f_type == 'Ground':
        pltovercurrent.SetAttribute("shownRelays", 3)                        # Ground
    else:
        pltovercurrent.SetAttribute("shownRelays", 1)                        # Phase
    # Reclose operation
    pltovercurrent.SetAttribute("recloserOperation", 0)                  # All
    # Cut curves at
    pltovercurrent.SetAttribute("cutCurves", 2)                         # Short-circuit/breaking current
    # Display results
    pltovercurrent.SetAttribute("showCalcResults", 1)                   # Currents
    # Consider breaker operating time
    pltovercurrent.SetAttribute("considerBreakerTime", 0)               # No
    # Show 'out of service' units
    pltovercurrent.SetAttribute("showOutOfServiceUnits", 0)             # No
    # Show grading margins while rag & drop
    pltovercurrent.SetAttribute("showGradingMarginDrag", 1)             # Yes
    # Grading margin +/-
    pltovercurrent.SetAttribute("gradingMarginDrag", 0.2)               # 0.2s
    # Points per curve
    pltovercurrent.SetAttribute("numPointsPerCurve", 200)               # 200


    # Text Format
    # =============================================================================

    # Number format
    # Format
    pltovercurrent.SetAttribute("numFormatMode", 1)                     # fixed decimals
    # Decimals
    pltovercurrent.SetAttribute("numFormatDecimals", 3)                 # 3
    # Exponent character
    pltovercurrent.SetAttribute("numExponentCase", 1)                   # "E"
    # Show signs for positive values
    pltovercurrent.SetAttribute("numFormatShowPlus", 0)                 # No

    # Label font
    # Label font
    # pltovercurrent.SetAttribute("editLabelFontButton", "Arial 10")      # Arial 10
    # Colour
    # pltovercurrent.SetAttribute("labelFontColor", 1)                    # Black?


    # Style and Layout
    # =============================================================================

    # Border and Background
    # Draw border
    pltovercurrent.SetAttribute("drawBorder", 0)                        # No
    # Border colour
    # pltovercurrent.SetAttribute("borderColor", 3)                     # Grey?
    # Border style
    # pltovercurrent.SetAttribute("borderStyle", 1)                     # Straight line
    # Border width
    #pltovercurrent.SetAttribute("borderWidth", "Hairline")             # Hairline
    # Fill background
    pltovercurrent.SetAttribute("fillBackground", 0)                    # No
    # Fill colour
    #pltovercurrent.SetAttribute("fillColor", 7)                        # White?

    # Curve area position (mm)
    # Auto-position Curve Area
    pltovercurrent.SetAttribute("autoPositionCurveArea", 1)             # Yes
    # Left
    # pltovercurrent.SetAttribute("curveAreaLeft", 0)                    # 0
    # Right
    # pltovercurrent.SetAttribute("curveAreaRight", 0)                   # 0
    # Bottom
    # pltovercurrent.SetAttribute("curveAreaBottom", 0)                  # 0
    # Top
    # pltovercurrent.SetAttribute("curveAreaTop", 0)                     # 0

    # Colour palette


def axis_settings(pltlinebarplot, f_type: str, devices: List[dd.Device]):
    """
    Configure TOC plot x- and y- axis extents
    """

    pltlinebarplot.SetAxisSharingLevelX(0)
    pltlinebarplot.SetAxisSharingLevelY(0)
    pltlinebarplot.SetScaleTypeX(1)
    pltlinebarplot.SetScaleTypeY(0)

    if f_type == 'Ground':
        x_min = _get_bound(devices[0].min_fl_pg, bound='Min')
        x_max = _get_bound(devices[0].max_fl_pg, bound='Max')
    else:
        max_fl_ph = max(devices[0].max_fl_2ph, devices[0].max_fl_3ph)
        x_min = _get_bound(devices[0].min_fl_2ph, bound='Min')
        x_max = _get_bound(max_fl_ph, bound='Max')
    pltlinebarplot.SetScaleX(x_min, x_max)
    pltlinebarplot.SetScaleY(0, 10)


def title_settings(pltlinebarplot, plot_name: str) -> None:
    """
    Configure TOC title bar
    """

    plttitle = pltlinebarplot.GetTitleObject()
    # Show title
    plttitle.SetAttribute("showTitle", 1)                                   # Yes
    # Title
    plttitle.SetAttribute("titleString", [plot_name])                      # name


def xvalue_settings(
    constant: pft.VisXvalue,
    name: str,
    value: float
) -> None:
    """
    Configure fault current marker line settings.

    Args:
        constant: VisXvalue object to configure.
        name: Label text for the marker.
        value: Fault current value in Amperes.
    """
    constant.loc_name = name
    constant.label = 1
    constant.lab_text = [name]
    constant.show = 1                                   # Show with intersections
    constant.iopt_lab = 3                               # Label position
    constant.value = value
    constant.color = 1
    constant.width = 5
    constant.xis = 0                                    # Current axi


def _get_bound(num: float, bound: str) -> float:
    """
    Round fault level to nearest order of magnitude.

    Args:
        num: Fault level value.
        bound: 'Min' to round down, 'Max' to round up.

    Returns:
        Rounded value for axis limit.
    """
    order_of_mag = 10 ** int(math.log10(num))

    if bound == 'Min':
        return math.floor(int(math.log10(num)))
        # return math.floor(num / order_of_mag) * order_of_mag
    else:
        return math.ceil(int(math.log10(num)))
        # return math.ceil(num / order_of_mag) * order_of_mag