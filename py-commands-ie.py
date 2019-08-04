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
        "value": value,
        "hover": hover,
        "wait": wait,
    }


def validate_project_parameters(parameters):
    
    traces = parameters["traces"]
    window = parameters["window"]
    find_element_dict = parameters["find_element_dict"]
    command = parameters["command"]
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
    elif command.upper() == "DOUBLECLICK":
        command = command.upper()
    elif command.upper() == "RIGHTCLICK":
        command = command.upper()
    elif command.upper() == "SEND":
        command = command.upper()
    elif command.upper() == "SELECT":
        command = command.upper()
    elif command.upper() == "LOCATION":
        command = command.upper()
    elif command.upper() == "WAIT":
        command = command.upper()
    else:
        return "ERROR: Invalid command parameter! Parameter = " + str(command)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tCommand = " + str(command))

    if "SEND" in command.upper() or "SELECT" in command.upper():
        if value == "":
            return "ERROR: Empty value parameters!"

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

    if command.upper() == "WAIT" and not wait.isdigit():
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
        "value": value,
        "hover": hover,
        "wait": wait,
    }
    
    
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


def find_element(ie_obj, find_element_dict, traces):
    
    while ie_obj.Busy:
        time.sleep(0.05)
        
    while ie_obj.ReadyState != 4:
        time.sleep(0.05) 
    
    while ie_obj.Document.ReadyState != "complete":
        time.sleep(0.05)
    
    element_obj = ie_obj.Document.Body

    try:
        for loop_x in range(0, len(find_element_dict)):
            temp_list = find_element_dict[str(loop_x + 1)].split(",")
    
            id = None
            name = None
            value = None
            title = None
            class_name = None
            tag_name = None
            inner_text = None
            inner_html = None
            iframe = None
            item = None

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

                elif "class_name=" in str(temp_list[loop_y]).strip():
                    class_name = str(temp_list[loop_y]).strip()
                    class_name = class_name[len("class")+1:]

                elif "tag_name=" in str(temp_list[loop_y]).strip():
                    tag_name = str(temp_list[loop_y]).strip()
                    tag_name = tag_name[len("tag_name")+1:]

                elif "name=" in str(temp_list[loop_y]).strip():
                    name = str(temp_list[loop_y]).strip()
                    name = name[len("name")+1:]

                elif "inner_text=" in str(temp_list[loop_y]).strip():
                    inner_text = str(temp_list[loop_y]).strip()
                    inner_text = inner_text[len("inner_text")+1:]

                elif "inner_html=" in str(temp_list[loop_y]).strip():
                    inner_html = str(temp_list[loop_y]).strip()
                    inner_html = inner_html[len("inner_html")+1:]

                elif "iframe=" in str(temp_list[loop_y]).strip():
                    iframe = str(temp_list[loop_y]).strip()
                    iframe = iframe[len("iframe")+1:]

                elif "item=" in str(temp_list[loop_y]).strip():
                    item = str(temp_list[loop_y]).strip()
                    item = item[len("item")+1:]
                    
            if loop_x == 0:
                if iframe is not None:
                    element_obj = element_obj.all.item(iframe)
                    try:
                        element_obj = [x.contentDocument for x in element_obj]
                    except:
                        element_obj = element_obj.contentDocument
                        
    
            if not isinstance(element_obj, list):
                element_obj = [element_obj]
                
            element_obj_candidates = []
                
            for element_temp in element_obj:
                
                try:
                    if id is not None:
                        elements_temp = element_temp.all.item(id)
                        
                    elif name is not None:    
                        elements_temp = element_temp.all.item(name)
                    
                    elif class_name is not None:
                        elements_temp = element_temp.getElementsByClassName(class_name)
                        
                    elif tag_name is not None:
                        elements_temp = element_temp.getElementsByTagName(tag_name)
                        
                    else:    
                        elements_temp = element_temp.all
                except:
                    elements_temp = element_temp.all
                    
                try:
                    element_obj_candidates = element_obj_candidates + [x.contentDocument for x in elements_temp]
                except:
                    if not isinstance(elements_temp, list):
                        elements_temp = [elements_temp]
                    element_obj_candidates = element_obj_candidates + elements_temp
                    
            element_obj_matches = []
                    
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
                        
                    if class_name is not None:
                        if element_obj_candidate.classname != class_name:
                            continue
                        
                    if tag_name is not None:
                        if str(element_obj_candidate.tagname).upper() != tag_name.upper():
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
                
            if item is not None:
                item = int(item)
                print(len(element_obj_matches))
                element_obj_matches = [element_obj_matches[item-1]]
                
            element_obj = element_obj_matches
            
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
    command = parameters["command"]
    value = parameters["value"]
    hover = parameters["hover"]
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

    element_obj = find_element(ie_obj, find_element_dict, traces)

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
    value = parameters["value"]
    hover = parameters["hover"]
    wait = parameters["wait"]
    
    try:

        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Focusing on window and element start * ===")
        
        win32gui.SetForegroundWindow(ie_obj.hwnd)
        element_obj.focus()

        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Focusing on window element end * ===")
        
        if command.upper() == "LOCATION":
            command = "LOCATION"

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print location start * ===")

            print(element_obj.offsetLeft)
            ie_obj_position = win32gui.GetWindowRect(ie_obj.hwnd)
            ie_obj_position_top, ie_obj_position_left, ie_obj_position_bottom, ie_obj_position_right = ie_obj_position[1], ie_obj_position[0], ie_obj_position[3], ie_obj_position[2]
            print(ie_obj_position_top, ie_obj_position_left, ie_obj_position_bottom, ie_obj_position_right)

            element_obj_position = element_obj.getBoundingClientRect()
            element_obj_position_top, element_obj_position_left, element_obj_position_bottom, element_obj_position_right = element_obj_position.top, element_obj_position.left, element_obj_position.bottom, element_obj_position.right
            print(element_obj_position_top, element_obj_position_left, element_obj_position_bottom, element_obj_position_right)
            
            viewport_obj_position = ie_obj.Document.Body.getBoundingClientRect()
            viewport_obj_position_top, viewport_obj_position_left, viewport_obj_position_bottom, viewport_obj_position_right = viewport_obj_position.top, viewport_obj_position.left, viewport_obj_position.bottom, viewport_obj_position.right
            print(viewport_obj_position_top, viewport_obj_position_left, viewport_obj_position_bottom, viewport_obj_position_right)

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print location end * ===")

        elif "CLICK" in command.upper():
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform click command start * ===")

            mouse_location_x, mouse_location_y = pyautogui.position()

            if command.upper() == "CLICK":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to click...")

                element_obj.click()
                # pyautogui.click(position)

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tClick complete!")

            elif command.upper() == "DOUBLECLICK":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to double-click...")

                # pyautogui.doubleClick(position)

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tDouble-click complete!")

            elif command.upper() == "RIGHTCLICK":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to right-click...")

                # pyautogui.rightClick(position)

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tRight-click complete!")

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform click command end * ===")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Move mouse back to original position start * ===")

            if hover is False:
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to move mouse back to position...")

                pyautogui.moveTo(mouse_location_x, mouse_location_y)

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

            try:
                element_obj.text = value
            except:
                try:
                    element_obj.value = value
                except:
                    element_obj.innerText = value
            
            # element_obj.focus()
            # keyboard.write(value)

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tWriting values complete!")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform send command end * ===")

        elif "SELECT" in command.upper():
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform select command start * ===")

            if value.startswith("item=") and value[5:].isdigit():
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tValue is digit, therefore, convert to integer")

                value = int(int(value[5:]) - 1)

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to select item: " + str(value))

            if isinstance(value, int):
                element_obj.selectedIndex = value
            else:
                options = []
                
                for i in range(element_obj.options.length - 1):
                    options.append(str(element_obj.options(i).text.strip()))
                
                if value not in options:
                    return f"ERROR: There is no '{value} option in the element!"                        
                else:
                    element_obj.selectedIndex = options.index(value)

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSelecting value complete!")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform select command end * ===")
                
    except:
        ie_obj, element_obj = None, None
        
        print(traceback.format_exc())
        
        return "ERROR: Unexpected issue!"
    
    ie_obj, element_obj = None, None
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
