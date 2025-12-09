#     RSSwapTool - A script to swap in up to date or enhanced rolling stock
#     for older versions of stock in Train Simulator scenarios.
#     Copyright (C) 2022 James McKenzie jrmknz@yahoo.co.uk
#
#     This program is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     This program is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <https://www.gnu.org/licenses/>.

import xml.etree.ElementTree as ET
import sys
import os
import subprocess
import platform
import configparser
import webbrowser
import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from pathlib import Path
from pathlib import PureWindowsPath

# If you want to run this script on Linux you must configure the path to the wine executable. You need wine in order to
# run the serz.exe utility which converts between .bin and .xml scenario files.
# If you're not running this script on Linux this line can be ignored.
wine_executable = '/usr/bin/wine'

# Initialise the script, set the look and feel and get the configuration
version_number = '1.0.10'
version_date = '9 December 2025'
vehicle_list = []
railworks_path = ''
config = configparser.ConfigParser()
script_path = Path(os.path.abspath(os.path.dirname(sys.argv[0])))
path_to_config = script_path / 'config.ini'
config.read(path_to_config)
if config.has_option('RailWorks', 'path'):
    railworks_path = config.get('RailWorks', 'path')

filename = ''
# sg.LOOK_AND_FEEL_TABLE['Railish'] = {'BACKGROUND': '#00384F', 'TEXT': '#FFFFFF', 'INPUT': '#FFFFFF',
#                                     'TEXT_INPUT': '#000000', 'SCROLL': '#99CC99', 'BUTTON': ('#FFFFFF', '#002A3C'),
#                                     'PROGRESS': ('#31636d', '#002A3C'), 'BORDER': 2, 'SLIDER_DEPTH': 0,
#                                     'PROGRESS_DEPTH': 2, }
# sg.theme('Railish')

about1 = 'Tool for listing rolling stock in Train Simulator (Dovetail Games) scenarios, bundled with RSSwapTool to provide a standalone tool to examine scenarios and list rolling stock.'
about2 = 'Version ' + version_number + ' / ' + version_date
about3 = 'Copyright 2025 JR McKenzie (jrmknz@yahoo.co.uk) https://github.com/jrmckenzie/RSSwapTool'
about4 = 'This program is free software: you can redistribute it and / or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.'
about5 = 'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with this program.  If not, see https://www.gnu.org/licenses/.'

def create_mainbuttons_frame(container):
    global railworks_path
    frame = ttk.Frame(container)
    frame.columnconfigure(0, weight=1)
    ttk.Label(frame, text='RSReportTool v' + version_number, font=("Helvetica", 14)).grid(row=0, column=0,
                                                                                          columnspan=4, sticky=tk.W)
    ttk.Label(frame, text='Rolling stock report generator for existing scenarios.',
              font=("Helvetica", 10)).grid(row=1, column=0, columnspan=4, sticky=tk.W)
    ttk.Button(frame,
        text='Select scenario file to examine',
        command=callback_filebrowser,
    ).grid(row=2, column=0, columnspan=4, sticky=tk.W)
    ttk.Button(frame, text='Examine!', command=examine_pressed).grid(column=0, row=3, sticky=tk.W)
    ttk.Button(frame, text='About', command=about_pressed).grid(column=1, row=3, sticky=tk.W)
    ttk.Button(frame, text='RailWorks Folder', command=exit_pressed).grid(column=2, row=3, sticky=tk.W)
    ttk.Button(frame, text='Exit', command=exit_pressed).grid(column=3, row=3, sticky=tk.W)
    ttk.Label(frame, text='RailWorks folder:',
              font=("Helvetica", 10)).grid(column=0, row=5, columnspan=4, sticky=tk.W)
    ttk.Entry(frame, textvariable=rw_path, width=60).grid(column=0, row=6, columnspan=4, sticky=tk.W)
    rw_path.set(railworks_path)
    ttk.Label(frame, text='© 2025 JR McKenzie', font=('Helvetica 7')).grid(row=7, column=0, columnspan=4, sticky=tk.W)

    for widget in frame.winfo_children():
        widget.grid(padx=5, pady=5)
    return frame

def examine_pressed():
    global filename
    # The examine button has been pressed
    if len(filename) < 1:
        print('No scenario selected!')
    else:
        scenarioPath = Path(filename)
        inFile = scenarioPath
        cmd = railworks_path / Path('serz.exe')
        serz_output = ''
        vehicle_list = []
        if str(scenarioPath.suffix) == '.bin':
            # This is a bin file so we need to run serz.exe command to convert it to a readable .xml
            # intermediate file
            if not cmd.is_file():
                print('serz.exe could not be found in ' + str(railworks_path) + '. Is this definitely your '
                                                                                   'RailWorks folder?',
                         'This application will now exit.')
                sys.exit()
            inFile = scenarioPath.parent / Path(str(scenarioPath.stem) + '-railvehicle_examination_report.xml')
            if platform.system() == 'Windows':
                p1 = subprocess.Popen([str(cmd), str(PureWindowsPath(scenarioPath)), '/xml:' +
                                       str(PureWindowsPath(inFile))], stdout=subprocess.PIPE)
            else:
                try:
                    wine_executable
                except NameError:
                    wine_executable = '/usr/bin/wine'
                p1 = subprocess.Popen([wine_executable, str(cmd), 'z:' + str(PureWindowsPath(scenarioPath)), '/xml:' +
                                       'z:' + str(PureWindowsPath(inFile))], stdout=subprocess.PIPE)
            p1.wait()
            serz_output = 'serz.exe ' + p1.communicate()[0].decode('ascii')
            # Now the intermediate .xml has been created by serz.exe, read it in to this script and do the
            # processing
            tree = parse_xml(inFile)
            inFile.unlink()
            #if tree is False:
            #    continue
        else:
            tree = parse_xml(inFile)
            #if tree is False:
            #    continue
        scenario_properties = parse_properties_xml(scenarioPath.parent)
        html_report_file = scenarioPath.parent / Path(str(scenarioPath.stem) +
                                                      '-railvehicle_examination_report.html')
        convert_vlist_to_html_table(html_report_file, scenario_properties)
        html_report_status_text = 'Report listing all rail vehicles located in ' + str(html_report_file)
        #browser = sg.popup_yes_no(html_report_status_text,
        #                          'Do you want to open the report in your web browser now?')
        #if browser == 'Yes':
        webbrowser.open(html_report_file.as_uri())
        # re-initialise vehicle list
        vehicle_list = []


def about_pressed():
    global about
    print('About pressed')
    top=Toplevel(tkwin)
    top.geometry("516x400")
    top.title("About RSReportTool")
    Label(top, text=about1, font=('Helvetica 10'), wraplength=500,justify="left").pack(side=tk.TOP, padx=8, pady=8, anchor=tk.W)
    Label(top, text=about2, font=('Helvetica 10'), wraplength=500,justify="left").pack(side=tk.TOP, padx=8, pady=8, anchor=tk.W)
    Label(top, text=about3, font=('Helvetica 10'), wraplength=500,justify="left").pack(side=tk.TOP, padx=8, pady=8, anchor=tk.W)
    Label(top, text=about4, font=('Helvetica 10'), wraplength=500,justify="left").pack(side=tk.TOP, padx=8, pady=8, anchor=tk.W)
    Label(top, text=about5, font=('Helvetica 10'), wraplength=500,justify="left").pack(side=tk.TOP, padx=8, pady=8, anchor=tk.W)

def callback_filebrowser():
    global filename
    filename = select_file()

def exit_pressed():
    sys.exit()

def select_folder():
    global railworks_path
    directory = filedialog.askdirectory()
    print(directory)

def select_file():
    filetypes = (
        ('Scenario files', 'Scenario*.bin'),
        ('All files', '*.*')
    )

    scenariofile = fd.askopenfilename(
        title='Open a file',
        initialdir=railworks_path,
        filetypes=filetypes)
    return(scenariofile)


window_width = 375
window_height = 225

def parse_xml(xml_file):
    # Check we can open the file, parse it and find some rail vehicle consists in it before proceeding
    try:
        parser_tree = ET.parse(xml_file)
    except FileNotFoundError:
        print('Scenario file ' + str(Path(xml_file)) + ' not found.', 'Please try again.', title='Error')
        return False
    except ET.ParseError:
        print('The file you requested (' + str(Path(xml_file)) + ') could not be processed due to an XML parse '
            'error. Is it definitely a scenario file?', 'Please try again with another Scenario.bin or Scenario.xml '
            'file.', title='Error')
        return False
    ET.register_namespace("d", "http://www.kuju.com/TnT/2003/Delta")
    root = parser_tree.getroot()
    consists = root.findall('./Record/cConsist')
    if len(consists) == 0:
        print('The file you requested (' + str(Path(xml_file)) + ') does not appear to contain any rail vehicle '
            'consists. Is it definitely a scenario file?', 'Please try again with another Scenario.bin or Scenario.xml '
            'file.', title='Error')
        return False
    consist_nr = 0
    for citem in consists:
        # Find the service name of the consist, if there is one, otherwise call it a loose consist.
        service = citem.find('Driver/cDriver/ServiceName/Localisation-cUserLocalisedString/English')
        if service is None:
            service = 'Loose consist'
        else:
            service = service.text
        # Find if this consist is driven by the player
        playerdriver = citem.find('Driver/cDriver/PlayerDriver')
        if playerdriver is None:
            playerdriven = False
        elif playerdriver.text == '1':
            playerdriven = True
        else:
            playerdriven = False
        # Iterate through RailVehicles list of the consist
        for rvehicles in citem.findall('RailVehicles'):
            # Iterate through each RailVehicle in the consist
            for coentity in rvehicles.findall('cOwnedEntity'):
                provider = coentity.find(
                    'BlueprintID/iBlueprintLibrary-cAbsoluteBlueprintID/BlueprintSetID/iBlueprintLibrary'
                    '-cBlueprintSetID/Provider')
                product = coentity.find(
                    'BlueprintID/iBlueprintLibrary-cAbsoluteBlueprintID/BlueprintSetID/iBlueprintLibrary'
                    '-cBlueprintSetID/Product')
                blueprint = coentity.find('BlueprintID/iBlueprintLibrary-cAbsoluteBlueprintID/BlueprintID')
                name = coentity.find('Name')
                number = coentity.find('Component/*/UniqueNumber')
                loaded = coentity.find('Component/cCargoComponent/IsPreLoaded')
                vehicle_list.append(
                    [str(consist_nr), provider.text, product.text, blueprint.text, name.text, number.text, loaded.text,
                     service, playerdriven])
        consist_nr += 1
    # All necessary elements processed, now close progress bar window and return the new xml tree object
    return parser_tree


def route_parser(file):
    try:
        route_tree = ET.parse(file)
    except FileNotFoundError:
        return False
    except ET.ParseError:
        return False
    return route_tree


def scenario_props_parser(file):
    try:
        parser_tree = ET.parse(file)
    except FileNotFoundError:
        return False
    except ET.ParseError:
        return False
    return parser_tree


def parse_properties_xml(dir):
    # Check we can open the file, parse it and find some rail vehicle consists in it before proceeding
    xml_props = Path(dir) / 'ScenarioProperties.xml'
    xml_route = Path(dir).parent.parent / 'RouteProperties.xml'
    props_tree = scenario_props_parser(xml_props)
    route_tree = route_parser(xml_route)
    ET.register_namespace("d", "http://www.kuju.com/TnT/2003/Delta")
    if props_tree:
        root = props_tree.getroot()
        DN = root.find('./DisplayName/Localisation-cUserLocalisedString/English')
        Desc = root.find('./Description/Localisation-cUserLocalisedString/English')
        Brief = root.find('./Briefing/Localisation-cUserLocalisedString/English')
        Start = root.find('./StartLocation/Localisation-cUserLocalisedString/English')
        properties = [DN.text, Desc.text, Brief.text, Start.text, 'unknown - ' + str(xml_route) + ' not found']
        if route_tree:
            root = route_tree.getroot()
            RN = root.find('./DisplayName/Localisation-cUserLocalisedString/English')
            properties[4] = RN.text
        return properties
    return False


def convert_vlist_to_html_table(html_file_path, scenarioProps):
    htmhead = '''<html lang="en">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<title>Scenario rail vehicle and asset report</title>
<link href='https://fonts.googleapis.com/css?family=Roboto' rel='stylesheet'>
<style>
body,.dataframe {
    font-family: 'Roboto';font-size: 10pt;
}
tr.shaded_row {
    background-color: #cccccc;
}
td.missing {
    color: #bb2222;
    font-style: italic;
}
h1 {
    font-family: 'Roboto';font-size: 24pt;
    font-style: bold;
}
h2 {
    font-family: 'Roboto';font-size: 18pt;
    font-style: bold;
}
h3,thead {
    font-family: 'Roboto';font-size: 14pt;
    font-style: bold;
    border-style: none none solid none;
    border-width: 1px;
}
</style>\n</head>\n<body>\n'''
    htmrv = '<h1>Rail vehicle list</h1>\n<table border=\'1\' class=\'dataframe\'>\n  <thead>\n' \
            '    <tr style=\'text-align: right;\'>\n      <th>Consist</th>\n      <th>Provider</th>\n' \
            '      <th>Product</th>\n      <th>Blueprint</th>\n      <th>Name</th>\n      <th>Number</th>\n' \
            '      <th>Loaded</th>\n    </tr>\n  </thead>\n  <tbody>\n'
    unique_assets = []
    last_cons = -1
    col_no = 0
    for row in vehicle_list:
        if row[1:3] not in unique_assets:
            unique_assets.append(row[1:3])
        if int(row[0]) > last_cons:
            # start of a new consist - count how many vehicles are in this consist
            rowspan = (list(zip(*vehicle_list))[0]).count(row[0])
        else:
            rowspan = 0
        col_htm = ''
        row[3] = row[3].replace('.xml', '.bin')
        my_path = Path(railworks_path, 'Assets', row[1], row[2], row[3].replace('\\','/'))
        if os.path.isfile(str(my_path)):
            tdstyle = ''
        else:
            tdstyle = ' class="missing"'
        cname = '<i>' + str(row[7]) + '</i>'
        if row[8] is True:
            # Consist is driven by the player - make the name bold and append (Player driven)
            cname = '<b>' + cname + '</b> (Player driven)'
        for col in row[0:7]:
            col_no += 1
            if rowspan > 0 and col_no == 1:
                col_htm = col_htm + '      <td rowspan=' + str(rowspan) + '>' + cname + '</td>\n'
            elif col_no > 1:
                col_htm = col_htm + '      <td' + tdstyle + '>' + col + '</td>\n'
        col_no = 0
        if (int(row[0]) % 2) == 0:
            htmrv = htmrv + '    <tr>\n' + col_htm + '    </tr>\n'
        else:
            htmrv = htmrv + '    <tr class=\'shaded_row\'>\n' + col_htm + '    </tr>\n'
        last_cons = int(row[0])
    htmrv = htmrv + '  </tbody>\n</table>\n<h3>' + str(len(vehicle_list)) + ' vehicles in total in this scenario.</h3>'
    htmas = '\n<h1>List of rail vehicle assets used</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n' \
            '    <tr style=\"text-align: right;\">\n      <th>Provider</th>\n      <th>Product</th>\n    </tr>\n' \
            '  </thead>\n  <tbody>\n'
    unique_assets.sort(key=lambda x: (x[0], x[1]))
    for asset in unique_assets:
        htmas = htmas + '    <tr>\n      <td>' + asset[0] + '</td>\n      <td>' + asset[1] + '</td>\n    </tr>\n'
    htmas = htmas + '  </tbody>\n</table>\n'
    htp = ''
    if scenarioProps:
        htp = '\n<h1>Scenario properties</h1>\n<table border=\"1\" class=\"dataframe\" style=\"text-align: left;\">\n' \
                '    <tr>\n      <th>Title</th>\n      <td>' + str(scenarioProps[0]) + '</td>\n    </tr>\n' \
                '    <tr>\n      <th>Description</th>\n      <td>' + str(scenarioProps[1]) + '</td>\n    </tr>\n' \
                '    <tr>\n      <th>Briefing</th>\n      <td>' + str(scenarioProps[2]) + '</td>\n    </tr>\n' \
                '    <tr>\n      <th>Start From</th>\n      <td>' + str(scenarioProps[3]) + '</td>\n    </tr>\n' \
                '    <tr>\n      <th>Route</th>\n      <td>' + str(scenarioProps[4]) + '</td>\n    </tr>\n' \
                '  </table>\n'
    htm = htmhead + htp + htmas + htmrv + '</body>\n</html>\n'
    html_file_path.touch()
    html_file_path.write_text(htm)
    return True


if __name__ == "__main__":
    tkwin = tk.Tk()
    rw_path = tk.StringVar()
    tkwin.resizable(False, False)
    tkwin.title('RSReportTool - Rolling stock report tool')

    # get the screen dimension
    screen_width = tkwin.winfo_screenwidth()
    screen_height = tkwin.winfo_screenheight()

    # find the center point
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)

    # set the position of the window to the center of the screen
    tkwin.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    button_frame = create_mainbuttons_frame(tkwin)
    button_frame.grid(column=0, row=0)

    tkwin.mainloop()
