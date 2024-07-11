""" Templates for different studies
"""

from dataclasses import dataclass


@dataclass
class SCMethod:
    IEC60909 = 1
    Complete = 3


@dataclass
class FaultType:
    _3_Ph_Short_Circuit = '3psc'
    _2_Ph_Short_Circuit = '2psc'
    Single_Phase_to_Ground = 'spgf'
    _2_Phase_to_Ground = '2pgf'
    _3_Ph_Short_Circuit_Unbalanced = '3rst'


@dataclass
class Calculate:
    Maximum = 0
    Minimum = 1


@dataclass
class ProtTrippingCurrent:
    SubTransient: int = 0
    Transient: int = 1
    MixedMode: int = 2


@dataclass
class EvtShcType:
    _3_Ph_Short_Circuit = 0
    _2_Ph_Short_Circuit = 1
    Single_Phase_to_Ground = 2
    _2_Phase_to_Ground = 3


@dataclass
class Max3PhaseShortCircuit:
    iopt_mde: int = SCMethod.Complete
    iopt_shc: str = FaultType._3_Ph_Short_Circuit
    iopt_cur: int = Calculate.Maximum
    iopt_cnf: bool = False
    ildfinit: bool = False
    cfac_full: float = 1.1
    iIgnLoad: bool = True
    iIgnLneCap: bool = True
    iIgnShnt: bool = True
    iopt_prot: int = 1
    iIksForProt: int = ProtTrippingCurrent.Transient
    Rf: float = 0
    Xf: float = 0


@dataclass
class Max2PhaseShortCircuit:
    iopt_mde: int = SCMethod.Complete
    iopt_shc: str = FaultType._2_Ph_Short_Circuit
    iopt_cur: int = Calculate.Maximum
    iopt_cnf: bool = False
    ildfinit: bool = False
    cfac_full: float = 1.1
    iIgnLoad: bool = True
    iIgnLneCap: bool = True
    iIgnShnt: bool = True
    iopt_prot: int = 1
    iIksForProt: int = ProtTrippingCurrent.Transient
    Rf: float = 0
    Xf: float = 0


@dataclass
class Min3PhaseShortCircuit:
    iopt_mde: int = SCMethod.Complete
    iopt_shc: str = FaultType._3_Ph_Short_Circuit
    iopt_cur: int = Calculate.Minimum
    i_p2psc: int = 0
    iopt_cnf: bool = False
    ildfinit: bool = False
    cfac_full: float = 1.0
    iIgnLoad: bool = True
    iIgnLneCap: bool = True
    iIgnShnt: bool = True
    iopt_prot: int = 1
    iIksForProt: int = ProtTrippingCurrent.Transient
    Rf: float = 0
    Xf: float = 0


@dataclass
class Min2PhaseShortCircuit:
    iopt_mde: int = SCMethod.Complete
    iopt_shc: str = FaultType._2_Ph_Short_Circuit
    iopt_cur: int = Calculate.Minimum
    i_p2psc: int = 0
    iopt_cnf: bool = False
    ildfinit: bool = False
    cfac_full: float = 1.0
    iIgnLoad: bool = True
    iIgnLneCap: bool = True
    iIgnShnt: bool = True
    iopt_prot: int = 1
    iIksForProt: int = ProtTrippingCurrent.Transient
    Rf: float = 0
    Xf: float = 0


@dataclass
class MaxGroundShortCircuit:
    iopt_mde: int = SCMethod.Complete
    iopt_shc: str = FaultType.Single_Phase_to_Ground
    iopt_cur: int = Calculate.Maximum
    i_pspgf: int = 0
    iopt_cnf: bool = False
    ildfinit: bool = False
    cfac_full: float = 1.1
    iIgnLoad: bool = True
    iIgnLneCap: bool = True
    iIgnShnt: bool = True
    iopt_prot: int = 1
    iIksForProt: int = ProtTrippingCurrent.Transient
    Rf: float = 0
    Xf: float = 0


@dataclass
class MinGroundShortCircuit:
    iopt_mde: int = SCMethod.Complete
    iopt_shc: str = FaultType.Single_Phase_to_Ground
    iopt_cur: int = Calculate.Minimum
    i_pspgf: int = 0
    iopt_cnf: bool = False
    ildfinit: bool = False
    cfac_full: float = 1.0
    iIgnLoad: bool = True
    iIgnLneCap: bool = True
    iIgnShnt: bool = True
    iopt_prot: int = 1
    iIksForProt: int = ProtTrippingCurrent.Transient
    Rf: float = 0
    Xf: float = 0


def apply_sc(ComShc, Format, Type):
    """
    Set all the attributes in the ComSch module to the values in the
    required dataclass
    :param ComShc: ComShc command method
    :param Format: 'Max', 'Min'
    :param Type: string '3 Phase', '2 Phase', 'Ground'
    :return: ComShc command method
    """

    if Format == 'Max':
        if Type == '3 Phase':
            for field in Max3PhaseShortCircuit.__dataclass_fields__:
                ComShc.SetAttribute(field, getattr(Max3PhaseShortCircuit, field))
        elif Type == '2 Phase':
            for field in Max2PhaseShortCircuit.__dataclass_fields__:
                ComShc.SetAttribute(field, getattr(Max2PhaseShortCircuit, field))
        else:
            for field in MaxGroundShortCircuit.__dataclass_fields__:
                ComShc.SetAttribute(field, getattr(MaxGroundShortCircuit, field))
    else:
        if Type == '3 Phase':
            for field in Min3PhaseShortCircuit.__dataclass_fields__:
                ComShc.SetAttribute(field, getattr(Min3PhaseShortCircuit, field))
        elif Type == '2 Phase':
            for field in Min2PhaseShortCircuit.__dataclass_fields__:
                ComShc.SetAttribute(field, getattr(Min2PhaseShortCircuit, field))
        else:
            for field in MinGroundShortCircuit.__dataclass_fields__:
                ComShc.SetAttribute(field, getattr(MinGroundShortCircuit, field))


@dataclass
class LoadFlowType:
    Balanced = 0
    Unbalanced = 1


@dataclass
class UnbalancedLoadFlow:
    iopt_net = LoadFlowType.Unbalanced


def apply_lf(ComLdf, Balanced=False):
    """applies the settintgs from the unbalanced load flow dataclass
    """
    if Balanced:
        pass
    else:
        for field in UnbalancedLoadFlow.__dataclass_fields__:
            ComLdf.SetAttribute(field, getattr(UnbalancedLoadFlow, field))