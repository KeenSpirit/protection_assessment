from devices import relays
from importlib import reload
import functools
import time
import logging
reload(relays)



def timer(func):
    """Print the runtime of the decorated function"""

    @functools.wraps(func)
    def wrapper_timer(*args, **kwargs):
        start_time = time.perf_counter()
        value = func(*args, **kwargs)
        end_time = time.perf_counter()
        run_time = end_time - start_time
        logging.info(f"Finished {func.__name__!r} in {run_time:.4f} secs")
        return value

    return wrapper_timer


def test_annotations(app):

    plot_name = 'OC Cape side'

    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    plot_page = graphics_board.GetContents(f"{plot_name}.GrpPage")[0]
    layers = plot_page.GetContents("Layers.IntFolder")[0]
    annotationLayer = layers.GetContents("Annotations.IntGrflayer")[0]
    dplMap = layers.GetContents("DPL Map.IntDplmap")[0]

    al.add_annot_layer(app, annotationLayer, dplMap)

    # dplMap.Clear()
    # dplMap.Insert('type','Line')
    # dplMap.Insert('color','{0,255,0}')
    # dplMap.Insert('points','{{10,100},{30,200}}')
    # dplMap.Insert('thickness','1')
    #
    # annotationLayer.AddAnnotationElement(dplMap)
    # annotationLayer.UpdateAnnotationElement(1, dplMap)



def test_reclosing(app, test_relay_name: str):
    """
    This function shows which protection elements are active according to their logic and the reclose sequence.
    Uer must determine whether the printed output meets expectations.
    These functions are used in the conductor damage algorithm to ensure correct worst case energy is assessed
    for each trip
    fault_type: 'Phase-Ground', '2-Phase'
    :param app:
    :return:
    """

    app.ClearOutputWindow()

    devices = get_devices(app)
    test_relay = [device for device in devices if device.loc_name == test_relay_name][0]
    app.PrintPlain(f"test relay:{test_relay}")
    fault_type = 'Phase-Ground'
    app.PrintPlain(f"fault_type: {fault_type}")
    total_trips = relays.get_device_trips(test_relay)
    app.PrintPlain(f"total_trips:{total_trips}")
    relays.reset_reclosing(test_relay)

    trip_count = 1
    while trip_count <= total_trips:
        app.PrintPlain(f"test trip_count:{trip_count}")
        block_service_status = relays.set_enabled_elements(test_relay)
        elements = relays.get_prot_elements(test_relay)
        active_elements = relays.get_active_elements(elements, fault_type)
        for element in active_elements:
            app.PrintPlain(f"active_element:{element}")
        relays.reset_block_service_status(block_service_status)
        trip_count = relays.trip_count(test_relay, increment=True)


def get_devices(app):
    net_mod = app.GetProjectFolder("netmod")
    # Filter for relays under network model recursively.
    all_relays = net_mod.GetContents("*.ElmRelay", True)
    relays = [
        relay
        for relay in all_relays
        if relay.cpGrid
        if relay.cpGrid.IsCalcRelevant()
        if relay.GetParent().GetClassName() == "StaCubic"
        if relay.fold_id.cterm.IsEnergized()
        if not relay.IsOutOfService()
    ]
    # Create a list of active fuses
    all_fuses = net_mod.GetContents("*.RelFuse", True)
    fuses = [
        fuse
        for fuse in all_fuses
        if fuse.cpGrid
        if fuse.cpGrid.IsCalcRelevant()
        if fuse.fold_id.HasAttribute("cterm")
        if fuse.fold_id.cterm.IsEnergized()
        if not fuse.IsOutOfService()
        if determine_fuse_type(fuse)
    ]
    devices = relays + fuses
    return devices


def determine_fuse_type(fuse):
    """This function will observe the fuse location and determine if it is
    a Distribution transformer fuse, SWER isolating fuse or a line fuse"""
    # First check is that if the fuse exists in a terminal that is in the
    # System Overiew then it will be a line fuse.
    fuse_active = fuse.HasAttribute("r:fold_id:r:obj_id:e:loc_name")
    if not fuse_active:
        return True
    fuse_grid = fuse.cpGrid
    if (
            fuse.GetAttribute("r:fold_id:r:cterm:r:fold_id:e:loc_name")
            == fuse_grid.loc_name
    ):
        # This would indicate it is in a line cubical
        return True
    if fuse.loc_name not in fuse.GetAttribute("r:fold_id:r:obj_id:e:loc_name"):
        # This indicates that the fuse is not in a switch object
        return True
    secondary_sub = fuse.fold_id.cterm.fold_id
    contents = secondary_sub.GetContents()
    for content in contents:
        if content.GetClassName() == "ElmTr2":
            return False
    else:
        return True
