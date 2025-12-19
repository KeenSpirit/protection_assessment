""" This module holds code to support protection script. The functions
supported in here are:
    App -> A contect manager for the powerfactory application
    **
"""

import uuid
from contextlib import contextmanager
import sys
from pf_config import pft

__all__ = ['app_manager'
    , 'project_manager'
    , 'temporary_variation'
    , 'obtain_region'
           ]


@contextmanager
def app_manager(app, clear=True, gui=False, echo_on=False, cache=False):
    """ Allows the powerfactor application class to be called using a with
    with statement.

    DONT USE cache=True unless you really know the impacts, this will hurt you.
    """

    try:
        app.ResetCalculation()
        if clear: app.ClearOutputWindow()
        if echo_on:
            app.EchoOn()
        else:
            echo = app.GetFromStudyCase('ComEcho')
            echo.iopt_err = True
            echo.iopt_wrng = False
            echo.iopt_info = False
            echo.iopt_oth = True
            app.EchoOff()

        app.SetGuiUpdateEnabled(1) if gui else app.SetGuiUpdateEnabled(0)
        app.SetWriteCacheEnabled(1) if cache else app.SetWriteCacheEnabled(0)
        app.SetUserBreakEnabled(1)
        yield app

    except:
        raise

    finally:
        app.EchoOn()
        app.SetGuiUpdateEnabled(1)
        app.SetUserBreakEnabled(0)
        if app.IsWriteCacheEnabled():
            app.WriteChangesToDb()
            app.SetWriteCacheEnabled(0)
        app.ClearRecycleBin()
        del app


@contextmanager
def project_manager(app):

    try:
        # new folder under Project/Library/Equipment Type library/
        temporary_library = app.GetLocalLibrary().CreateObject('IntFolder', 'Temp Types')
        yield temporary_library

    except:
        raise

    finally:
        temporary_library.Delete()


@contextmanager
def temporary_variation(app):
    variation_name = str(uuid.uuid1())
    variation_time = app.GetActiveStudyCase().GetAttribute('iStudyTime')
    net_dat = app.GetProjectFolder("netmod")
    variation_folder = net_dat.GetContents("Variations")[0]

    try:
        variation = variation_folder.CreateObject("IntScheme", variation_name)
        variation.Activate()
        variation.NewStage(variation_name, variation_time, 1)
        yield variation

    except:
        raise

    finally:
        variation.Deactivate()
        variation.Delete()


def obtain_region(app):

    project = app.GetActiveProject()
    derived_proj = project.der_baseproject
    der_proj_name = derived_proj.GetFullName()

    regional_model = 'Regional Models'
    seq_model = 'SEQ'

    if regional_model in der_proj_name:
        # This is a regional model
        region=regional_model
    elif seq_model in der_proj_name:
        # This is a SEQ model
        region = seq_model
    else:
        msg = (
            "The appropriate region for the model could not be found. "
            "Please contact the script administrator to resolve this issue."
        )
        raise RuntimeError(msg)
    return region


def create_obj(parent, obj_name: str, obj_class: str):

    obj = parent.GetContents(f"{obj_name}.{obj_class}")

    if not obj:
        obj = parent.CreateObject(obj_class, obj_name)
    else:
        obj = obj[0]
    return obj