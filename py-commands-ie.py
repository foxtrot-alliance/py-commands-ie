import sys
import time
import datetime
import traceback
import pyautogui
import keyboard
import win32com.client as win32
import win32gui


def retrieve_project_parameters():
    
    parameters = sys.argv

    parameters_number = parameters.index("-traces") if "-traces" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        traces = parameters[parameters_number]
    else:
        traces = ""

    parameters_number = parameters.index("-window") if "-window" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        window = parameters[parameters_number]
    else:
        window = ""

    find_element_dict = {}

    parameters_number = parameters.index("-find_element1") if "-find_element1" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        find_element_dict = {"1": parameters[parameters_number]}

    parameters_number = parameters.index("-find_element2") if "-find_element2" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        find_element_dict["2"] = parameters[parameters_number]

    parameters_number = parameters.index("-find_element3") if "-find_element3" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        find_element_dict["3"] = parameters[parameters_number]

    parameters_number = parameters.index("-find_element4") if "-find_element4" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        find_element_dict["4"] = parameters[parameters_number]

    parameters_number = parameters.index("-find_element5") if "-find_element5" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        find_element_dict["5"] = parameters[parameters_number]

    parameters_number = parameters.index("-command") if "-command" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        command = parameters[parameters_number]
    else:
        command = ""

    parameters_number = parameters.index("-attribute") if "-attribute" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        attribute = parameters[parameters_number]
    else:
        attribute = ""

    parameters_number = parameters.index("-value") if "-value" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        value = parameters[parameters_number]
    else:
        value = ""

    parameters_number = parameters.index("-hover") if "-hover" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        hover = parameters[parameters_number]
    else:
        hover = ""

    parameters_number = parameters.index("-wait") if "-wait" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        wait = parameters[parameters_number]
    else:
        wait = ""
        
    return {
        "traces": traces,
        "window": window,
        "find_element_dict": find_element_dict,
        "command": command,
        "attribute": attribute,
        "value": value,
        "hover": hover,
        "wait": wait,
    }


def validate_project_parameters(parameters):
    
    traces = parameters["traces"]
    window = parameters["window"]
    find_element_dict = parameters["find_element_dict"]
    command = parameters["command"]
    attribute = parameters["attribute"]
    value = parameters["value"]
    hover = parameters["hover"]
    wait = parameters["wait"]
    
    if traces == "" or traces.upper() == "FALSE":
        traces = False
    elif traces.upper() == "TRUE":
        traces = True
    else:
        return "ERROR: Invalid traces parameter! Parameter = " + str(traces)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Parameters retrieved start * ===")

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tWindow = " + str(window))

    for x in range(0, len(find_element_dict)):
        if find_element_dict[str(x + 1)] == "":
            return "ERROR: Empty find element parameters!"

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tFind element(s) = " + str(find_element_dict))

    if command.upper() == "CLICK":
        command = command.upper()
    elif command.upper() == "CLICK_BYPASS":
        command = command.upper()
    elif command.upper() == "DOUBLECLICK":
        command = command.upper()
    elif command.upper() == "RIGHTCLICK":
        command = command.upper()
    elif command.upper() == "SEND":
        command = command.upper()
    elif command.upper() == "SET":
        command = command.upper()
    elif command.upper() == "GET":
        command = command.upper()
    elif command.upper() == "SELECT":
        command = command.upper()
    elif command.upper() == "COUNT":
        command = command.upper()
    elif command.upper() == "LOCATION":
        command = command.upper()
    else:
        return "ERROR: Invalid command parameter! Parameter = " + str(command)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tCommand = " + str(command))

    if "SEND" in command.upper() or "SELECT" in command.upper():
        if value == "":
            return "ERROR: Empty value parameters!"

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttribute = " + str(attribute))

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tValue = " + str(value))

    if hover == "" or hover.upper() == "FALSE":
        hover = False
    elif hover.upper() == "TRUE":
        hover = True
    else:
        return "ERROR: Invalid hover parameter! Parameter = " + str(hover)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tHover = " + str(hover))

    if wait == "" or wait.upper() == "TRUE":
        wait = True
    elif wait.upper() == "FALSE":
        wait = False
    else:
        return "ERROR: Invalid wait parameter! Parameter = " + str(wait)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tWait = " + str(wait))

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Parameters retrieved end * ===")
        
    return {
        "traces": traces,
        "window": window,
        "find_element_dict": find_element_dict,
        "command": command,
        "attribute": attribute,
        "value": value,
        "hover": hover,
        "wait": wait,
    }
    
def get_element_location(ie_obj, element_obj, traces):
    
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tCalculating element location...")

    ie_obj_position = win32gui.GetWindowRect(ie_obj.hwnd)
    ie_obj_position_top, ie_obj_position_left, ie_obj_position_bottom, ie_obj_position_right = ie_obj_position[1], ie_obj_position[0], ie_obj_position[3], ie_obj_position[2]
    
    ie_obj_position_left = ie_obj_position_left + 7
    ie_obj_position_bottom = ie_obj_position_bottom - 8
    ie_obj_position_right = ie_obj_position_right - 8
    ie_obj_position_top = ie_obj_position_top + ((ie_obj_position_bottom - ie_obj_position_top) - ie_obj.Document.documentElement.clientHeight)
    
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tINTERNET EXPLORER = Left: {str(ie_obj_position_left)}, Top: {str(ie_obj_position_top)}, Right: {str(ie_obj_position_right)}, Bottom: {str(ie_obj_position_bottom)}")

    element_obj_position = element_obj.getBoundingClientRect()
    element_obj_position_top, element_obj_position_left, element_obj_position_bottom, element_obj_position_right = element_obj_position.top, element_obj_position.left, element_obj_position.bottom, element_obj_position.right
    
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tELEMENT = Left: {str(element_obj_position_left)}, Top: {str(element_obj_position_top)}, Right: {str(element_obj_position_right)}, Bottom: {str(element_obj_position_bottom)}")
    
    left = ie_obj_position_left + element_obj_position_left
    top = ie_obj_position_top + element_obj_position_top
    right = ie_obj_position_left + element_obj_position_right
    bottom = ie_obj_position_top + element_obj_position_bottom
    
    x = left + ((right-left) / 2)
    y = top + ((bottom-top) / 2)
    
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tCALCULATED LOCATION = X: {str(x)}, Y: {str(y)}")
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tLocation calculated!")
    
    ie_obj, element_obj = None, None
    
    return x, y
    
def find_window(window, traces):
    
    try:
        ie_obj = None
        ie_obj_windows = []
        app_obj = win32.Dispatch("Shell.Application")

        for window_obj in app_obj.Windows():
            if "Internet Explorer" in str(window_obj.Name):
                ie_obj_windows.append(window_obj)
                
        if window == "" or window is None:
            ie_obj = ie_obj_windows[0]
            
        elif window.isdigit():
            try:
                ie_obj = ie_obj_windows[int(window)]
            except:
                pass
            
        else:
            for window_obj in ie_obj_windows:
                if window in str(window_obj.Name):
                    ie_obj = window_obj
                    break
    except:
        print(traceback.format_exc())
        ie_obj = None
    
    finally:
        app_obj = None
        ie_obj_windows = None
        return ie_obj


def find_element(ie_obj, find_element_dict, wait, traces):
    
    while ie_obj.Busy:
        time.sleep(0.05)
    
    if wait is not False:
        while ie_obj.ReadyState != 4:
            time.sleep(0.05) 
        
        while ie_obj.Document.ReadyState != "complete":
            time.sleep(0.05)
    
    element_obj = ie_obj.Document

    try:
        for loop_x in range(0, len(find_element_dict)):
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tLocating element {loop_x + 1}...")
                
            id = None
            name = None
            value = None
            title = None
            classname = None
            tagname = None
            inner_text = None
            inner_html = None
            parent = "False"
            iframe = None
            item = None
    
            temp_list = find_element_dict[str(loop_x + 1)].split(",")
                    
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\tParameters = {temp_list}")

            for loop_y in range(0, len(temp_list)):
                if "id=" in str(temp_list[loop_y]).strip():
                    id = str(temp_list[loop_y]).strip()
                    id = id[len("id")+1:]

                elif "value=" in str(temp_list[loop_y]).strip():
                    value = str(temp_list[loop_y]).strip()
                    value = value[len("value")+1:]

                elif "title=" in str(temp_list[loop_y]).strip():
                    title = str(temp_list[loop_y]).strip()
                    title = title[len("title")+1:]

                elif "classname=" in str(temp_list[loop_y]).strip():
                    classname = str(temp_list[loop_y]).strip()
                    classname = classname[len("classname")+1:]

                elif "tagname=" in str(temp_list[loop_y]).strip():
                    tagname = str(temp_list[loop_y]).strip()
                    tagname = tagname[len("tagname")+1:]

                elif "name=" in str(temp_list[loop_y]).strip():
                    name = str(temp_list[loop_y]).strip()
                    name = name[len("name")+1:]

                elif "innertext=" in str(temp_list[loop_y]).strip():
                    innertext = str(temp_list[loop_y]).strip()
                    innertext = innertext[len("innertext")+1:]

                elif "innerhtml=" in str(temp_list[loop_y]).strip():
                    innerhtml = str(temp_list[loop_y]).strip()
                    innerhtml = innerhtml[len("innerhtml")+1:]

                elif "parent=" in str(temp_list[loop_y]).strip():
                    parent = str(temp_list[loop_y]).strip()
                    parent = parent[len("parent")+1:]

                elif "iframe=" in str(temp_list[loop_y]).strip():
                    iframe = str(temp_list[loop_y]).strip()
                    iframe = iframe[len("iframe")+1:]

                elif "item=" in str(temp_list[loop_y]).strip():
                    item = str(temp_list[loop_y]).strip()
                    item = item[len("item")+1:]
            
            if loop_x == 0 and iframe is not None:
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\tAttempting to connect to iframe...")
                    
                if isinstance(element_obj, list):
                    element_obj = element_obj[0]
                
                try:
                    element_obj = element_obj.getElementById(iframe)
                except:
                    try:
                        element_obj = element_obj.body.getElementById(iframe)
                    except:
                        element_obj = element_obj.all.item(iframe)
                
                try:
                    element_obj = [x.contentDocument for x in element_obj]
                except:
                    if isinstance(element_obj, list):
                        element_obj = element_obj[0]
                        
                    element_obj = element_obj.contentDocument
                    
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\tConnected to iframe!")

            if not isinstance(element_obj, list):
                element_obj = [element_obj]
                
            element_obj_candidates = []
                
            for element_index, element_temp in enumerate(element_obj):
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\tLocating subelement {element_index + 1}...")
                    
                try:
                    id_applied = False
                    name_applied = False
                    classname_applied = False
                    tagname_applied = False
                    
                    if parent.upper() == "TRUE":
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tDetecting parent element...")
                            
                        elements_temp = element_temp.parentElement
                        
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tParent element detected!")
                    
                    elif id is not None:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatching on ID...")
                            
                        try:
                            elements_temp = element_temp.getElementById(id)
                            
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.getElementById(id)'!")
                                
                        except:
                            try:
                                elements_temp = element_temp.body.getElementById(id)
                            
                                if traces is True:
                                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.body.getElementById(id)'!")
                                    
                            except:
                                elements_temp = element_temp.all.item(id)
                            
                                if traces is True:
                                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.all.item(id)'!")
                            
                        id_applied = True
                        
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatched on ID!")
                        
                    elif name is not None:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatching on name...")
                            
                        try:
                            elements_temp = element_temp.getElementsByName(name)
                            
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.getElementsByName(name)'!")
                                
                        except:
                            try:
                                elements_temp = element_temp.body.getElementsByName(name)
                            
                                if traces is True:
                                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.body.getElementsByName(name)'!")
                                    
                            except:
                                elements_temp = element_temp.all.item(name)
                            
                                if traces is True:
                                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.all.item(name)'!")
                                
                        name_applied = True
                        
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatched on name!")
                    
                    elif classname is not None:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatching on classname...")
                            
                        try:
                            elements_temp = element_temp.getElementsByClassName(classname)
                            
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.getElementsByClassName(classname)'!")
                                
                        except:
                            elements_temp = element_temp.body.getElementsByClassName(classname)
                            
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.body.getElementsByClassName(classname)'!")
                            
                        classname_applied = True
                        
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatched on classname!")
                        
                    elif tagname is not None:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatching on tagname...")
                            
                        try:
                            elements_temp = element_temp.getElementsByTagName(tagname)
                            
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.getElementsByTagName(tagname)'!")
                                
                        except:
                            try:
                                elements_temp = element_temp.body.getElementsByTagName(tagname)
                            
                                if traces is True:
                                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.body.getElementsByTagName(tagname)'!")
                                    
                            except:
                                elements_temp = element_temp.all.tags(tagname)
                            
                                if traces is True:
                                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tSuccesfully using '.all.tags(tagname)'!")
                                
                        tagname_applied = True
                        
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatched on tagname!")
                        
                    else:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatching all...")
                            
                        elements_temp = element_temp.all
                        
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatched all!")
                        
                except:
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tAttempt to match failed, matching all instead...")
                    
                    elements_temp = element_temp.all
                        
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatched all!")
                
                try:
                    element_obj_candidates = element_obj_candidates + [x for x in elements_temp]
                    
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tValidating matches...")
                    
                    if id is not None and id_applied is True:
                        if [x.id for x in elements_temp][0] != id:
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tMatches invalid, reversing to all!")
                                
                            elements_temp = element_temp.all
                            element_obj_candidates = element_obj_candidates + [x for x in elements_temp]
                    
                    if name is not None and name_applied is True:
                        if [x.name for x in elements_temp][0] != name:
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tMatches invalid, reversing to all!")
                                
                            elements_temp = element_temp.all
                            element_obj_candidates = element_obj_candidates + [x for x in elements_temp]
                    
                    if classname is not None and classname_applied is True:
                        if [x.classname for x in elements_temp][0] != classname:
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tMatches invalid, reversing to all!")
                                
                            elements_temp = element_temp.all
                            element_obj_candidates = element_obj_candidates + [x for x in elements_temp]
                    
                    if tagname is not None and tagname_applied is True:
                        if [x.tagname for x in elements_temp][0] != tagname.upper():
                            if traces is True:
                                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\t\tMatches invalid, reversing to all!")
                                
                            elements_temp = element_temp.all
                            element_obj_candidates = element_obj_candidates + [x for x in elements_temp]
                    
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\t\tMatches validated!")
                    
                except:                            
                    if not isinstance(elements_temp, list):
                        elements_temp = [elements_temp]
                    element_obj_candidates = element_obj_candidates + elements_temp
                    
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\tSubelement {element_index + 1} handled!")
                    
            element_obj_matches = []
                    
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\tChecking all {len(element_obj_candidates)} subelement candidates...")
                    
            for element_obj_candidate in element_obj_candidates:
                try:
                    if id is not None:
                        if element_obj_candidate.id != id:
                            continue
                        
                    if value is not None:
                        if element_obj_candidate.value != value:
                            continue
                        
                    if title is not None:
                        if element_obj_candidate.title != title:
                            continue
                        
                    if name is not None:
                        if element_obj_candidate.name != name:
                            continue
                        
                    if classname is not None:
                        if element_obj_candidate.classname != classname:
                            continue
                        
                    if tagname is not None:
                        if str(element_obj_candidate.tagname).upper() != tagname.upper():
                            continue
                        
                    if inner_text is not None:
                        if str(element_obj_candidate.innertext) != inner_text:
                            continue
                        
                    if inner_html is not None:
                        if str(element_obj_candidate.innerhtml) != inner_html:
                            continue
                        
                except:
                    continue
                
                element_obj_matches.append(element_obj_candidate)
                    
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\t\tAll {len(element_obj_candidates)} subelement candidates checked, {len(element_obj_matches)} OK!")
                
            if item is not None:
                item = int(item)
                element_obj_matches = [element_obj_matches[item-1]]
                
            element_obj = element_obj_matches
            
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tElement {loop_x + 1} located!")
            
    except:
        element_obj = []
        print(traceback.format_exc())
        
    finally:
        ie_obj = None
        element_obj_matches = None
        element_obj_candidates = None
        elements_temp = None
        
        if len(element_obj) == 0:
            return None
        elif len(element_obj) == 1:
            return element_obj[0]
        else:
            return element_obj


def locate_target(parameters):
    
    traces = parameters["traces"]
    window = parameters["window"]
    find_element_dict = parameters["find_element_dict"]
    wait = parameters["wait"]
    
    ie_obj = None
    element_obj = None
    
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Locate the target start * ===")
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tLocating the Internet Explorer application window...")

    ie_obj = find_window(window, traces)

    if ie_obj is None:
        return ie_obj, element_obj, "ERROR: Internet Explorer application window not located!"
    else:
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tInternet Explorer application window located!")

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tLocating the element(s)...")

    element_obj = find_element(ie_obj, find_element_dict, wait, traces)

    if element_obj is None:
        return ie_obj, element_obj, "ERROR: Element(s) not located!"
    else:
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tElement(s) located!")

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Locate the target end * ===")
        
    return ie_obj, element_obj, "SUCCESS"


def execute_command(parameters, ie_obj, element_obj):
    
    traces = parameters["traces"]
    window = parameters["window"]
    find_element_dict = parameters["find_element_dict"]
    command = parameters["command"]
    attribute = parameters["attribute"]
    value = parameters["value"]
    hover = parameters["hover"]
    wait = parameters["wait"]
    
    try:
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Focusing on window and element start * ===")
        
        if not "GET" in command.upper() and not "COUNT" in command.upper():
            win32gui.SetForegroundWindow(ie_obj.hwnd)
                
        if not isinstance(element_obj, list):
            element_objs = [element_obj]
        else:
            element_objs = element_obj
            
        if command.upper() == "COUNT":
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print element count start * ===")
                
            print(len(element_objs))
            
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print element count end * ===")
                
            ie_obj, element_obj, element_objs = None, None, None
            return True
            
        for element_obj in element_objs:
            while ie_obj.Busy:
                time.sleep(0.05)
            
            if not "GET" in command.upper() and not "COUNT" in command.upper():
                element_obj.focus()

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Focusing on window element end * ===")
            
            if command.upper() == "LOCATION":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print location start * ===")
                
                x, y = get_element_location(ie_obj, element_obj, traces)
                print(f"X={x}, Y={y}")

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print location end * ===")

            elif "CLICK" in command.upper():
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform click command start * ===")

                if command.upper() == "CLICK_BYPASS":
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to click...")

                    element_obj.click()

                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tClick complete!")
                    
                else:
                    mouse_location_x, mouse_location_y = pyautogui.position()
                    element_location_x, element_location_y = get_element_location(ie_obj, element_obj, traces)

                    if command.upper() == "CLICK":
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to click...")
                    
                        pyautogui.click(x=element_location_x, y=element_location_y)

                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tClick complete!")

                    elif command.upper() == "DOUBLECLICK":
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to double-click...")
                    
                        pyautogui.doubleClick(x=element_location_x, y=element_location_y)

                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tDouble-click complete!")

                    elif command.upper() == "RIGHTCLICK":
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to right-click...")
                    
                        pyautogui.rightClick(x=element_location_x, y=element_location_y)

                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tRight-click complete!")

                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform click command end * ===")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Move mouse back to original position start * ===")

                    if hover is False:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to move mouse back to position...")

                        pyautogui.moveTo(x=mouse_location_x, y=mouse_location_y)

                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tMoving mouse back to position complete!")

                    else:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tHovering mouse activated, therefore, do NOT move mouse back in position")

                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Move mouse back to original position end * ===")

            elif "SEND" in command.upper():
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform send command start * ===")
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to write values: " + str(value))
                
                keyboard.write(value)

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tWriting values complete!")
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform send command end * ===")
                    
            elif "SET" in command.upper():
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform set command start * ===")
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tAttempting to set the attribute '{attribute}' to the value: " + str(value))
                    
                if attribute.upper() == "TEXT":
                    element_obj.text = value
                elif attribute.upper() == "VALUE":
                    element_obj.value = value
                elif attribute.upper() == "INNERTEXT":
                    element_obj.innerText = value
                elif attribute.upper() == "INNERHTML":
                    element_obj.innerHtml = value
                else:
                    element_obj.setAttribute(attribute, value)

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSetting the attribute to the value complete!")
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform set command end * ===")
                    
            elif "GET" in command.upper():
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform get command start * ===")
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + f"\tAttempting to get the attribute '{attribute}'")
                    
                if attribute.upper() == "TEXT":
                    try:
                        print(str(element_obj.text).strip())
                    except:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAs bytes, using unicode...")
                            
                        print(element_obj.text.encode('utf-8').strip())
                        
                elif attribute.upper() == "VALUE":
                    try:
                        print(str(element_obj.value).strip())
                    except:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAs bytes, using unicode...")
                            
                        print(element_obj.value.encode('utf-8').strip())
                        
                elif attribute.upper() == "INNERTEXT":
                    try:
                        print(str(element_obj.innerText).strip())
                    except:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAs bytes, using unicode...")
                            
                        print(element_obj.innerText.encode('utf-8').strip())
                        
                elif attribute.upper() == "INNERHTML":
                    try:
                        print(str(element_obj.innerHtml).strip())
                    except:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAs bytes, using unicode...")
                            
                        print(element_obj.innerHtml.encode('utf-8').strip())
                        
                else:
                    try:
                        print(str(element_obj.getAttribute(attribute)).strip())
                    except:
                        if traces is True:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAs bytes, using unicode...")
                            
                        print(element_obj.getAttribute(attribute).encode('utf-8').strip())

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tGetting the attribute complete!")
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform get command end * ===")

            elif "SELECT" in command.upper():
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform select command start * ===")

                if str(value).startswith("item=") and str(value)[5:].isdigit():
                    if traces is True:
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tValue is digit, therefore, convert to integer")

                    value = int(int(str(value)[5:]) - 1)

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to select item: " + str(value))

                if isinstance(value, int):
                    element_obj.selectedIndex = value
                else:
                    options_text = []
                    options_value = []
                    
                    for option in element_obj.options:
                        try:
                            options_text.append(str(option.text.strip()))
                        except:
                            options_text.append("ERROR")
                        
                        try:
                            options_value.append(str(option.value.strip()))
                        except:
                            options_value.append("ERROR")
                    
                    if value not in options_text and value not in options_value:
                        ie_obj, element_obj, element_objs = None, None, None
                        return f"ERROR: There is no '{value}' option in the element!"
                    
                    elif value in options_text:
                        element_obj.selectedIndex = options_text.index(value)
                        
                    elif value in options_value:
                        element_obj.selectedIndex = options_value.index(value)
                        
                try:
                    element_obj.FireEvent("onchange")
                except:
                    pass

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSelecting value complete!")
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform select command end * ===")
                    
    except:
        print(traceback.format_exc())
        
        ie_obj, element_obj, element_objs = None, None, None
        return "ERROR: Unexpected issue!"
    
    ie_obj, element_obj, element_objs = None, None, None
    return True


def main():
    
    parameters = retrieve_project_parameters()
    
    parameters = validate_project_parameters(parameters)
    if not isinstance(parameters, dict):
        print(str(parameters))
        return
    
    ie_obj, element_obj, valid = locate_target(parameters)
    if "ERROR" in valid:
        print(str(valid))
        ie_obj, element_obj = None, None
        return
    
    valid = execute_command(parameters, ie_obj, element_obj)
    if not valid is True:
        print(str(valid))
        ie_obj, element_obj = None, None
        return
    
    ie_obj, element_obj = None, None
    print("SUCCESS")
    
    
if __name__ == "__main__":
    main()
