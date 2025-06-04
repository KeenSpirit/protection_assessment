
def set_up(app):


    prjt = app.GetActiveProject()
    setting_folder = quick_filter_conf(app, prjt)
    clear_dpl_attr(app, prjt)
    # Configuration - conditional formatting for colour map.
    colour_condition_conf(app, setting_folder)


def quick_filter_conf(app, prjt):
    """
    PROJECT COLOUR SETTINGS: Provide the parameters and formulae for the colour
    filters.  This will also provide a quick way of filtering to find elements
    that meet a certain condition.
    """
    # Get the contents of the 'Project Colour Settings' folder.
    setting_folder = prjt.GetContents("*.SetFold", True)[0]
    # Define the variable element to - 'line'.  This will select line elements.
    name_1 = "Phase Flt Cond Damage"
    name_2 = "Earth Flt Cond Damage"
    elements = ["*.ElmLne"]

    # Using the list of filters provided in the main function, a for loop is
    # used to populate the 'Project Colour Settings' folder with the required
    # colour filters, removing any existing filters with the same name.
    existing_general_filters = setting_folder.GetContents("*.IntFilt", True)
    for general_filter in existing_general_filters:
        if name_1 in general_filter.loc_name or name_2 in general_filter.loc_name:
            general_filter.Delete()
    general_filter_1 = setting_folder.CreateObject("IntFilt", name_1)
    general_filter_2 = setting_folder.CreateObject("IntFilt", name_2)
    # format the colour filters with the regional filter requirements
    conditional_config(general_filter_1, elements)
    # format the colour filters with the regional filter requirements
    conditional_config(general_filter_2, elements)

    # Return the variable 'setting_folder' it will be needed for the colour
    # configuration function
    return setting_folder


def conditional_config(obj, elements):
    """
    Based upon the selected region, this function will configure the
    conditional formatting required for both the colour condition and quick
    filter functions.
    """
    # For each study type, create a list of conditions for the selected region.
    list_of_cond_names = ["Damaged", "Undamaged", "SWER", "No Data"]

    # Check 'obj.loc_name' to determine if the scenario is reach or damage.
    # For reach scenarios, get the appropriate region's configuration otherwise
    # get the damage configuration.
    if "Phase Flt Cond Damage" in obj.loc_name:
        dpl_num = "dpl1"
        damage_condition_config(list_of_cond_names, obj, elements, dpl_num)
    if "Earth Flt Cond Damage" in obj.loc_name:
        dpl_num = "dpl2"
        damage_condition_config(list_of_cond_names, obj, elements, dpl_num)


def damage_condition_config(list_of_cond_names, obj, elements, dpl_num):
    """The energy of the conductor is referenced against its rating. These
    conditions will display where the ratings are exceeded."""

    for cond_name in list_of_cond_names:
        obj.CreateObject("SetFilt", cond_name)
        condition = obj.GetContents(f"{cond_name}.SetFilt")[0]
        condition.SetAttribute("objset", elements)
        condition.SetAttribute("icalcrel", 1)
        condition.SetAttribute("icoups", 0)
        if "Damaged" in cond_name:
            condition.SetAttribute("expr", [f"e:{dpl_num}=1"])
            condition.SetAttribute("color", 2)
        elif "Undamaged" in cond_name:
            condition.SetAttribute("expr", [f"e:{dpl_num}=2"])
            condition.SetAttribute("color", 3)
        elif "SWER" in cond_name:
            condition.SetAttribute("expr", [f"e:{dpl_num}=3"])
            condition.SetAttribute("color", 6)
        else:
            condition.SetAttribute("expr", [f"e:{dpl_num}=0"])
            condition.SetAttribute("color", 9)


def clear_dpl_attr(app, prjt):
    """This function will clear all the dpl attributes of elements in the model
    """
    app.SetGraphicUpdate(0)
    # Create a list of elements that could have had results written to them.
    active_lines = [
        line
        for line in prjt.GetContents("*.ElmLne", True)
        if line.GetAttribute("cpGrid")
    ]
    # Loop through all elements and only write to attributes that have got
    # results already written too.
    for element in active_lines:
        if element.GetAttribute("e:dpl1"):
            element.SetAttribute("e:dpl1", 0)
        if element.GetAttribute("e:dpl2"):
            element.SetAttribute("e:dpl2", 0)
        if element.GetAttribute("e:dpl3"):
            element.SetAttribute("e:dpl3", 0)
        if element.GetAttribute("e:dpl4"):
            element.SetAttribute("e:dpl4", 0)
        if element.GetAttribute("e:dpl5"):
            element.SetAttribute("e:dpl5", 0)
    app.SetGraphicUpdate(1)


def colour_condition_conf(app, setting_folder):
    """
    This function will create the conditional formatting objects under the
    'Settings' tab in the project folder
    """
    # Get the contents of the Project 'Colour Settings folder'
    # Define the List of elements that the conditional formatting will be
    # applied to.
    names = ["Phase Flt Cond Damage", "Earth Flt Cond Damage"]
    project_colour_fold = setting_folder.GetContents("*.SetColours", True)[0]

    # Define the variable element to - 'line'.  This will select line elements.
    elements = ["*.ElmLne"]

    # For loop to step through the list of filter names
    for name in names:
        existing_sets_of_cond = project_colour_fold.GetContents("*.IntFiltSet", True)
        # If an object already exists then it is to be deleted so that the
        # latest setting can be applied.
        for set_of_con in existing_sets_of_cond:
            if name in set_of_con.loc_name:
                set_of_con.Delete()
        set_of_conditions = project_colour_fold.CreateObject("IntFiltset", name)
        conditional_config(set_of_conditions, elements)
    return