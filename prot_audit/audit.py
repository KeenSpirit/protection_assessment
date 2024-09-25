from pf_protection_helper import create_obj

def audit_all_relays(app, relay_list):

    for relay in relay_list:
        if relay.object.GetClassName() == 'ElmRelay':
            relay_name = relay.object.loc_name
            name = f"{relay_name} Protection Audit"
            sect_devices = [device.object for device in relay.sect_lines]
            net_area = set_select(app, name, sect_devices)
            app.PrintPlain(f"Performing protection audit on {relay_name}")
            if relay.min_fl_ph:
                prot_audit(app, net_area, phase_faults=True)
            else:
                prot_audit(app, net_area)


def prot_audit(app, net_area, phase_faults=False):

    Protaudit = app.GetFromStudyCase("Protection Audit.ComProtaudit")

    # Network area
    Protaudit.p_sel = net_area

    # Calulation commands
    # Load Flow (Protaudit.cpLdf)
    ComLdf = app.GetFromStudyCase("Load Flow Calculation.ComLdf")
    #  Short-Circuit Sweer Short-Circuit Command
    # Fault types and impedances are not defined here
    ComShc = Protaudit.cpShc
    # Method complete
    ComShc.iopt_mde = 3
    # Minimum Short-Circuit Currents
    ComShc.iopt_cur = 1

    # Fault case definitions
    ComShcsweep = Protaudit.GetContents("Short-Circuit Sweep.ComShcsweep")[0]
    IntEvtshc = create_obj(ComShcsweep, "Short Circuits", "IntEvtshc")
    Events = IntEvtshc.GetContents()
    if Events: [Event.Delete() for Event in Events]
    if phase_faults:
        phase_EvtShc = IntEvtshc.CreateObject("EvtShc", "Phase Fault Event")
        phase_EvtShc.SetAttribute('i_shc', 1)
        phase_EvtShc.SetAttribute('i_p2psc', 1)           # a-b phase
        phase_EvtShc.SetAttribute('ZfaultInp', 0)
        phase_EvtShc.SetAttribute('R_f', 0)
        phase_EvtShc.SetAttribute('X_f', 0)
    ground_EvtShc = IntEvtshc.CreateObject("EvtShc", "Ground Fault Event")
    ground_EvtShc.SetAttribute('i_shc', 2)
    ground_EvtShc.SetAttribute('i_pspgf', 0)            # a phase
    ground_EvtShc.SetAttribute('ZfaultInp', 0)
    ground_EvtShc.SetAttribute('R_f', 0)
    ground_EvtShc.SetAttribute('X_f', 0)

    # Considered network equipment
    Protaudit.branches = 1
    Protaudit.busbars = 0
    Protaudit.stepSize = 50
    Protaudit.minLength = 0

    # Results
    study_case = app.GetActiveStudyCase()
    results_name = f"{net_area.loc_name} Results"
    results_file = study_case.GetContents(f"{results_name}.ElmRes")
    if results_file: [folder.Delete() for folder in results_file]
    results_file = study_case.CreateObject("ElmRes", results_name)
    Protaudit.p_res = results_file

    # Reporting
    Protaudit.iReport = 1
    ComAuditreport = Protaudit.cpOutputCmd
    audit_reporting(ComAuditreport)

    Protaudit.Execute()

    # prjt = app.GetActiveProject()
    # study_folder = prjt.GetContents("Protection Coordination Studies.IntFolder")[0]
    # study_folder.AddCopy(results_file, results_name)
    # results_file.Delete()


def set_select(app, name: str, elements: list):

    study_case = app.GetActiveStudyCase()
    new_set = study_case.CreateObject("SetSelect", name)
    new_set.iused = 4               # Short-Circuit Sweep
    new_set.iusedSub = 0            # Locations for short-circuit sweep

    new_set.AddRef(elements)
    return new_set


def audit_reporting(ComAuditreport):

    ComAuditreport.iopt_coord = 1
    ComAuditreport.iopt_trip = 1
    ComAuditreport.iopt_fct = 0
    ComAuditreport.iopt_style = 0

    ComAuditreport.useBreaker = 0
    ComAuditreport.coordMargin = 0.3

    ComAuditreport.coordSev = [0.0, 3.0]            # ["Failure", "No issue"]
    ComAuditreport.trip3ph = [1, 3, 3]
    ComAuditreport.trip2ph = [1, 3, 3]
    ComAuditreport.trip1ph = [1, 3, 3]
    ComAuditreport.trip3phSev = [0.0, 0.0, 3.0]     # ["Failure", "Failure", "No issue"]
    ComAuditreport.trip2phSev = [0.0, 0.0, 3.0]     # ["Failure", "Failure", "No issue"]
    ComAuditreport.trip1phSev = [0.0, 0.0, 3.0]     # ["Failure", "Failure", "No issue"]









