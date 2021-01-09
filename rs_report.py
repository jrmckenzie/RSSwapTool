#     RSSwapTool - A script to swap in up to date or enhanced rolling stock
#     for older versions of stock in Train Simulator scenarios.
#     Copyright (C) 2021 James McKenzie jrmckenzie @ gmail . com
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
import subprocess
import configparser
import PySimpleGUI as sg
from pathlib import Path

swaps_db = {}
report = []
rv_pairs = []
vehicle_list = []

def parse_xml(xml_file):
    try:
        parser_tree = ET.parse(xml_file)
    except FileNotFoundError:
        sys.exit('Scenario file ' + str(Path(xml_file)) + ' not found.', 'Please try again.', title='Error')
        return False
    except ET.ParseError:
        sys.exit('Scenario file ' + str(Path(xml_file)) + ' not found.', 'Please try again.', title='Error')
        return False
    ET.register_namespace("d", "http://www.kuju.com/TnT/2003/Delta")
    root = parser_tree.getroot()
    # iterate through the consists
    for citem in root.findall('./Record/cConsist'):
        # iterate through railvehicles list of the consist
        for rvehicles in citem.findall('RailVehicles'):
            # iterate through each railvehicle in the consist
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
                vehicle_list.append([provider.text, product.text, blueprint.text, name.text, number.text, loaded.text])
        for driver_inrvs in citem.findall('Driver/cDriver/InitialRV'):
            # iterate through driver instructions and update changed vehicle numbers in the consist
            for drv in driver_inrvs.findall('e'):
                for rvp in rv_pairs:
                    if drv.text == rvp[0]:
                        drv.text = rvp[1]
    for citem in root.findall('./Record/cConsist'):
        for cons_rvs in citem.findall(
                'Driver/cDriver/DriverInstructionContainer/cDriverInstructionContainer/DriverInstruction/'
                'cConsistOperations/DeltaTarget/cDriverInstructionTarget/RailVehicleNumber'):
            # iterate through driver consist instructions and update changed vehicle numbers in the consist
            for crv in cons_rvs.findall('e'):
                for rvp in rv_pairs:
                    if crv.text == rvp[0]:
                        crv.text = rvp[1]
    # All necessary elements processed, now return the new xml scenario to be written
    return parser_tree

tree = parse_xml('scenario.xml')
print(report)

def convert_vlist_to_html_table(html_file_path):
    htmhead = '''<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<title>Scenario rail vehicle and asset report</title>
<link href='https://fonts.googleapis.com/css?family=Roboto' rel='stylesheet'>
<style>
body,.dataframe {
    font-family: 'Roboto';font-size: 10pt;
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
    htmrv = "<h1>Rail vehicle list</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n" \
            "    <tr style=\"text-align: right;\">\n      <th>Provider</th>\n      <th>Product</th>\n" \
            "      <th>Blueprint</th>\n      <th>Name</th>\n      <th>Number</th>\n      <th>Loaded</th>\n    </tr>\n" \
            "  </thead>\n  <tbody>\n"
    unique_assets = []
    for row in vehicle_list:
        col_htm = ""
        if row[0:2] not in unique_assets:
            unique_assets.append(row[0:2])
        for col in row:
            col_htm = col_htm + "      <td>" + col + "</td>\n"
        htmrv = htmrv + "    <tr>\n" + col_htm + "    </tr>\n"
    htmrv = htmrv + "  </tbody>\n</table>\n<h3>" + str(len(vehicle_list)) + ' vehicles in total in this scenario.</h3>'
    htmas = "\n<h1>List of rail vehicle assets used</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n" \
            "    <tr style=\"text-align: right;\">\n      <th>Provider</th>\n      <th>Product</th>\n    </tr>\n" \
            "  </thead>\n  <tbody>\n"
    unique_assets.sort(key=lambda x: (x[0], x[1]))
    for asset in unique_assets:
        htmas = htmas + "    <tr>\n      <td>" + asset[0] + "</td>\n      <td>" + asset[1] + "</td>\n    </tr>\n"
    htmas = htmas + "  </tbody>\n</table>\n"
    htm = htmhead + htmas + htmrv + "</body>\n</html>\n"
    html_file_path.touch()
    html_file_path.write_text(htm)
    return True


railworks_path = ''
config = configparser.ConfigParser()
config.read('config.ini')

# Read configuration and find location of RailWorks folder, or ask user to set it
if config.has_option('RailWorks', 'path'):
    railworks_path = config.get('RailWorks', 'path')
else:
    loclayout = [[sg.T('')],
                 [sg.Text('Please locate your RailWorks folder:'), sg.Input(key='-IN2-', change_submits=False, readonly=True),
                  sg.FolderBrowse(key='RWloc')], [sg.Button('Submit')]]
    locwindow = sg.Window('Configure path to RailWorks folder', loclayout, size=(640, 150))
    while True:
        event, values = locwindow.read()
        if event == sg.WIN_CLOSED:
            if len(values['RWloc']) > 1:
                break
            else:
                sg.Popup('Please browse for the path to your RailWorks folder and try again.')
                continue
        elif event == 'Submit':
            if len(values['RWloc']) > 1:
                railworks_path = values['RWloc']
            else:
                sg.Popup('Please browse for the path to your RailWorks folder and try again.')
                continue
            if not config.has_section('RailWorks'):
                config.add_section('RailWorks')
            config.set('RailWorks', 'path', values['RWloc'])
            with open(Path('config.ini'), 'w') as iconfigfile:
                config.write(iconfigfile)
                iconfigfile.close()
            break
    locwindow.close()

# Set the layout of the GUI
main_column = [
    [sg.Text('RSSwapTool', font='Helvetica 16')],
    [sg.Text('Rolling stock report generator for existing scenarios.')],
    [sg.FileBrowse('Select scenario file to examine', key='Scenario_xml', tooltip='Locate the scenario .bin or .xml '
                                                                                  'file you wish to examine')],
    [sg.Button('Examine!'), sg.Button('Settings'), sg.Button('About'), sg.Button('Exit')],
    [sg.Text('© 2021 JR McKenzie', font='Helvetica 7')],
]

layout = [
    [
        sg.Column(main_column),
    ]
]

if __name__ == "__main__":
    window = sg.Window('RSSwapTool - Rolling stock swap tool', layout)
    while True:
        event, values = window.read()
        if event == 'Exit' or event == sg.WIN_CLOSED:
            break
        elif event == 'About':
            sg.Popup('About RSSwapTool',
                     'Tool for swapping rolling stock in Train Simulator (Dovetail Games) scenarios',
                     'Issued under the GNU General Public License - see https://www.gnu.org/licenses/',
                     'Version 0.1a',
                     'Copyright 2021 JR McKenzie', 'https://github.com/jrmckenzie/RSSwapTool')
        elif event == 'Settings':
            # The settings button has been pressed, so allow the user to change the RailWorks folder setting
            loclayout = [
                [sg.Text('Settings', justification='c')],
                [sg.Text('Path to RailWorks folder:'),
                 sg.Input(default_text=str(railworks_path), key='RWloc', readonly=True),
                 sg.FolderBrowse(key='RWloc')],
                [sg.Button('Save changes'), sg.Button('Cancel')]
            ]
            locwindow = sg.Window('Configure path to RailWorks folder', loclayout)
            while True:
                levent, lvalues = locwindow.read()
                if levent == 'Cancel' or sg.WIN_CLOSED:
                    break
                elif levent == 'Save changes':
                    railworks_path = lvalues['RWloc']
                    if not config.has_section('RailWorks'):
                        config.add_section('RailWorks')
                    config.set('RailWorks', 'path', lvalues['RWloc'])
                    with open(Path('config.ini'), 'w') as configfile:
                        config.write(configfile)
                        configfile.close()
                    railworks_path = Path(railworks_path)
                    break
            locwindow.close()
        elif event == 'Examine!':
            # The examine button has been pressed
            if len(values['Scenario_xml']) < 1:
                sg.popup('No scenario selected!')
            else:
                scenarioPath = Path(values['Scenario_xml'])
                inFile = scenarioPath
                cmd = railworks_path / Path('serz.exe')
                serz_output = ''
                if str(scenarioPath.suffix) == '.bin':
                    # This is a bin file so we need to run serz.exe command to convert it to a readable .xml
                    # intermediate file
                    inFile = scenarioPath.parent / Path(str(scenarioPath.stem) + '.xml')
                    p1 = subprocess.Popen([str(cmd), str(scenarioPath), '/xml:' + str(inFile)], stdout=subprocess.PIPE)
                    p1.wait()
                    serz_output = 'serz.exe ' + p1.communicate()[0].decode('ascii')
                    # Now the intermediate .xml has been created by serz.exe, read it in to this script and do the
                    # processing
                tree = parse_xml(inFile)
                html_report_file = scenarioPath.parent / Path(str(scenarioPath.stem) + '-railvehicle_report.html')
                convert_vlist_to_html_table(html_report_file)
                html_report_status_text = 'Report listing all rail vehicles located in ' + str(html_report_file)
                sg.popup(html_report_status_text)
