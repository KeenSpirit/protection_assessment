""" This module holds code to support protection script. The functions
supported in here are:
    App -> A contect manager for the powerfactory application
    **
"""

import uuid
from contextlib import contextmanager
import functools
import time

__all__ = ['app_manager'
    , 'project_manager'
    , 'temporary_variation'
    , 'debug'
    , 'timer'
           ]


def debug(app, func):
    """Print the function signature and return value"""

    @functools.wraps(func)
    def wrapper_debug(*args, **kwargs):
        args_repr = [repr(a) for a in args]
        kwargs_repr = [f"{k}={v!r}" for k, v in kwargs.items()]
        signature = ", ".join(args_repr + kwargs_repr)
        app.PrintPlain(f"Calling {func.__name__}({signature})")
        value = func(*args, **kwargs)
        app.PrintPlain(f"{func.__name__!r} returned {value!r}")
        return value

    return wrapper_debug


def timer(app, func):
    """Print the runtime of the decorated function"""

    @functools.wraps(func)
    def wrapper_timer(*args, **kwargs):
        start_time = time.perf_counter()
        value = func(*args, **kwargs)
        end_time = time.perf_counter()
        run_time = end_time - start_time
        app.PrintPlain(f"Finished {func.__name__!r} in {run_time:.4f} secs")
        return value

    return wrapper_timer


@contextmanager
def app_manager(app, clear=True, gui=False, echo=False, cache=False):
    """ Allows the powerfactor application class to be called using a with
    with statement.

    DONT USE cache=True unless you really know the impacts, this will hurt you.
    """

    try:
        app.ResetCalculation()
        if clear: app.ClearOutputWindow()

        if echo:
            app.EchoOn()
        else:
            echo.iopt_err = True
            echo.iopt_wrng = False
            echo.iopt_info = False
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
def project_manager(app, project=None):
    if project == None:
        project = app.GetActiveProject()

    try:
        # new folder under Project/Library/Equipment Type library/
        # TODO: Folder should have a unique name
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