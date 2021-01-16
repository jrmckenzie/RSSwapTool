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

import re
import random
import csv
import xml.etree.ElementTree as ET
import time
import sys
import os
import subprocess
import configparser
import PySimpleGUI as sg
import webbrowser
from pathlib import Path
from data_file import haa_e_wagons, haa_l_wagons, hto_e_wagons, hto_l_wagons, htv_e_wagons, htv_l_wagons, \
    vda_e_wagons, vda_l_wagons, HTO_141_numbers, HTO_143_numbers, HTO_146_numbers, HTO_rebodied_numbers, \
    HTV_146_numbers, HTV_rebodied_numbers, c158_s9bl_rr, c158_s9bl_nr, c158_s9bl_fgw, c158_s9bl_tpe, c158_s9bl_swt, \
    c158_nwc, c158_dtg_fc, c158_livman_rr, ap40headcodes_69_77, ap40headcodes_62_69

rv_list = []
rv_pairs = []
layout = []
vehicle_list = []
values = {}
vehicle_db = {}
user_db = {}
vp_blue_47_db = {}
railworks_path = ''
c56_opts = ['Use nearest numbered AP enhanced loco', 'Retain original loco if no matching AP plaque / sector available']
c86_opts = ['Use VP headcode blinds', 'Use AP plated box with markers', 'Do not swap this loco']
sg.LOOK_AND_FEEL_TABLE['Railish'] = {'BACKGROUND': '#00384F',
                                     'TEXT': '#FFFFFF',
                                     'INPUT': '#FFFFFF',
                                     'TEXT_INPUT': '#000000',
                                     'SCROLL': '#99CC99',
                                     'BUTTON': ('#FFFFFF', '#002A3C'),
                                     'PROGRESS': ('#31636d', '#002A3C'),
                                     'BORDER': 2, 'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 2, }
sg.theme('Railish')
config = configparser.ConfigParser()
path_to_config = Path(os.path.realpath(__file__)).parent / 'config.ini'
config.read(path_to_config)
# Read configuration and find location of RailWorks folder, or ask user to set it
if config.has_option('RailWorks', 'path'):
    railworks_path = config.get('RailWorks', 'path')
else:
    loclayout = [[sg.T('')],
                 [sg.Text('Please locate your RailWorks folder:'), sg.Input(key='-IN2-', change_submits=False,
                                                                            readonly=True),
                  sg.FolderBrowse(key='RWloc')], [sg.Button('Submit')]]
    locwindow = sg.Window('Configure path to RailWorks folder', loclayout, size=(640, 150))
    while True:
        event, values = locwindow.read()
        if event == sg.WIN_CLOSED:
            if values is not None:
                if values['RWloc'] is not None and len(values['RWloc']) > 1:
                    break
            else:
                sg.Popup('You must specify the path to your RailWorks folder for this application to work. '
                         'The application will now close.')
                sys.exit()
        elif event == 'Submit':
            if values['RWloc'] is not None and len(values['RWloc']) > 1:
                railworks_path = values['RWloc']
            else:
                sg.Popup('Please browse for the path to your RailWorks folder and try again.')
                continue
            if not config.has_section('RailWorks'):
                config.add_section('RailWorks')
            config.set('RailWorks', 'path', values['RWloc'])
            with open(path_to_config, 'w') as iconfigfile:
                config.write(iconfigfile)
                iconfigfile.close()
            break
    locwindow.close()
if not config.has_section('defaults'):
    config.add_section('defaults')
    config.set('defaults', 'replace_mk1', 'True')
    config.set('defaults', 'replace_mk2ac', 'True')
    config.set('defaults', 'replace_mk2df', 'True')
    config.set('defaults', 'replace_fsa', 'True')
    config.set('defaults', 'replace_haa', 'True')
    config.set('defaults', 'replace_hto', 'True')
    config.set('defaults', 'replace_htv', 'True')
    config.set('defaults', 'replace_vda', 'True')
    config.set('defaults', 'replace_ihh', 'False')
    config.set('defaults', 'replace_user', 'False')
    config.set('defaults', 'replace_c31', 'True')
    config.set('defaults', 'replace_c37', 'True')
    config.set('defaults', 'replace_c40', 'True')
    config.set('defaults', 'replace_c47', 'True')
    config.set('defaults', 'replace_c50', 'True')
    config.set('defaults', 'replace_c56', 'True')
    config.set('defaults', 'replace_c66', 'True')
    config.set('defaults', 'replace_c67', 'True')
    config.set('defaults', 'replace_c68', 'True')
    config.set('defaults', 'replace_c86', 'True')
    config.set('defaults', 'replace_hst', 'True')
    config.set('defaults', 'replace_c91', 'False')
    config.set('defaults', 'replace_c101', 'False')
    config.set('defaults', 'replace_c156', 'False')
    config.set('defaults', 'replace_c158', 'False')
    config.set('defaults', 'save_report', 'False')
    config.set('defaults', 'c56_rf', c56_opts[0])
    config.set('defaults', 'c86_hc', c86_opts[0])
    with open(path_to_config, 'w') as iconfigfile:
        config.write(iconfigfile)
        iconfigfile.close()


def import_data_from_csv(csv_filename):
    try:
        with open(Path(csv_filename), 'r') as csv_file:
            reader = csv.reader(csv_file)
            seen = ''
            for row in reader:
                outrow = []
                c = 0
                for col in row:
                    if c == 0:
                        key = col
                    elif c == 3:
                        outrow.append(re.escape(col))
                    else:
                        outrow.append(col)
                    c = c + 1
                if key == seen:
                    vehicle_db[key].append(outrow)
                else:
                    vehicle_db[key] = [outrow]
                seen = key
            return vehicle_db
    except FileNotFoundError:
        sg.popup('Error: vehicle swap data file ' + csv_filename + ' not found. Try re-installing the program.')
        sys.exit('Error: vehicle swap data file ' + csv_filename + ' not found. Try re-installing the program.')


# Read in the csv database of vehicles, substitutes and swap datal, store in VehicleDB dictionary
user_db_path = Path('tables/User.csv')
if not user_db_path.is_file():
    head = 'Label,Provider,Product,Blueprint,ReplaceProvider,ReplaceProduct,ReplaceBlueprint,ReplaceName,NumbersDcsv\n'
    user_db_path.touch()
    user_db_path.write_text(head)
vehicle_db = import_data_from_csv('tables/Replacements.csv')
user_db = import_data_from_csv('tables/User.csv')
vp_blue_47_db = import_data_from_csv('tables/Class47BRBlue_numbers.csv')

# Set the layout of the GUI
left_column = [
    [sg.Text('RSSwapTool\n', font='Helvetica 16')],
    [sg.FileBrowse('Select scenario file to process', key='Scenario_xml', tooltip='Locate the scenario .bin or .xml '
                                                                                  'file you wish to process')],
    [sg.Text('Tick the boxes below to choose the\nsubstitutions you would like to make.')],
    [sg.Checkbox('Replace Mk1 coaches', default=config.getboolean('defaults', 'replace_mk1'), enable_events=True,
                 tooltip='Tick to enable replacing of Mk1 coaches with AP Mk1 Coach Pack Vol. 1',
                 key='Replace_Mk1')],
    [sg.Checkbox('Replace Mk2A-C coaches', default=config.getboolean('defaults', 'replace_mk2ac'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Mk2a coaches with AP Mk2A-C Pack', key='Replace_Mk2ac')],
    [sg.Checkbox('Replace Mk2D-F coaches', default=config.getboolean('defaults', 'replace_mk2df'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Mk2e coaches with AP Mk2D-F Pack', key='Replace_Mk2df')],
    [sg.Checkbox('Replace FSA/FTA wagons', default=config.getboolean('defaults', 'replace_fsa'), enable_events=True,
                 tooltip='Tick to enable replacing of FSA wagons with AP FSA/FTA Wagon Pack', key='Replace_FSA')],
    [sg.Checkbox('Replace HAA wagons', default=config.getboolean('defaults', 'replace_haa'), enable_events=True,
                 tooltip='Tick to enable replacing of HAA wagons with AP MGR Wagon Pack', key='Replace_HAA')],
    [sg.Checkbox('Replace unfitted 21t coal wagons', default=config.getboolean('defaults', 'replace_hto'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of unfitted 21t coal wagons with Fastline Simulation HTO wagons',
                 key='Replace_HTO')],
    [sg.Checkbox('Replace fitted 21t coal wagons', default=config.getboolean('defaults', 'replace_htv'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of fitted 21t coal wagons with Fastline Simulation HTV wagons',
                 key='Replace_HTV')],
    [sg.Checkbox('Replace VDA wagons', default=config.getboolean('defaults', 'replace_vda'), enable_events=True,
                 tooltip='Tick to enable replacing of VDA wagons with Fastline Simulation VDA pack',
                 key='Replace_VDA')],
    [sg.Checkbox('Replace IHH stock', default=config.getboolean('defaults', 'replace_ihh'), enable_events=True,
                 tooltip='Tick to enable replacing of old Iron Horse House (IHH) stock, if your scenario contains any'
                         ' (if in doubt, leave this unticked)',
                 key='Replace_IHH')],
    [sg.Checkbox('Replace User-configured stock', default=config.getboolean('defaults', 'replace_user'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of user-configured stock, contained in file User.csv '
                         '(leave this unticked unless you have added your own substitutions to User.csv).',
                 key='Replace_User')],
    [sg.Text('© 2021 JR McKenzie', font='Helvetica 7')],
]
right_column = [
    [sg.Checkbox('Replace Class 31s', default=config.getboolean('defaults', 'replace_c31'), enable_events=True,
                 tooltip='Replace Class 31s with AP enhancement pack equivalent', key='Replace_C31')],
    [sg.Checkbox('Replace Class 37s', default=config.getboolean('defaults', 'replace_c37'), enable_events=True,
                 tooltip='Replace Class 37s with AP equivalent', key='Replace_C37')],
    [sg.Checkbox('Replace Class 40s', default=config.getboolean('defaults', 'replace_c40'), enable_events=True,
                 tooltip='Replace DT Class 40s with AP/RailRight equivalent', key='Replace_C40')],
    [sg.Checkbox('Replace Class 47s', default=config.getboolean('defaults', 'replace_c47'), enable_events=True,
                 tooltip='Replace BR Blue Class 47s with Vulcan Productions BR Blue Class 47 Pack versions',
                 key='Replace_C47')],
    [sg.Checkbox('Replace Class 50s', default=config.getboolean('defaults', 'replace_c50'), enable_events=True,
                 tooltip='Replace MeshTools Class 50s with AP equivalent', key='Replace_C50')],
    [sg.Checkbox('Replace Class 56s', default=config.getboolean('defaults', 'replace_c56'), enable_events=True,
                 tooltip='Replace RSC Class 56 Railfreight Sectors with AP enhancement pack equivalent',
                 key='Replace_C56')],
    [sg.Checkbox('Replace Class 66s', default=config.getboolean('defaults', 'replace_c66'), enable_events=True,
                 tooltip='Replace Class 66s with AP enhancement pack equivalent', key='Replace_C66')],
    [sg.Checkbox('Replace Class 67s', default=config.getboolean('defaults', 'replace_c67'), enable_events=True,
                 tooltip='Replace Class 67s with AP enhancement pack equivalent', key='Replace_C67')],
    [sg.Checkbox('Replace Class 68s', default=config.getboolean('defaults', 'replace_c68'), enable_events=True,
                 tooltip='Replace Class 68s with AP enhancement pack equivalent', key='Replace_C68')],
    [sg.Checkbox('Replace Class 86s', default=config.getboolean('defaults', 'replace_c86'), enable_events=True,
                 tooltip='Replace Class 86s with AP enhancement pack equivalent', key='Replace_C86')],
    [sg.Checkbox('Replace HST sets', default=config.getboolean('defaults', 'replace_hst'), enable_events=True,
                 tooltip='Tick to enable replacing of HST sets with AP enhanced versions (Valenta, MTU, VP185)',
                 key='Replace_HST')],
    [sg.Checkbox('Replace Class 91 EC sets', default=config.getboolean('defaults', 'replace_c91'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 91 East Coast sets with AP enhanced versions',
                 key='Replace_C91')],
    [sg.Checkbox('Replace Class 101 sets', default=config.getboolean('defaults', 'replace_c101'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of retired RSC Class101Pack with RSC BritishRailClass101 sets',
                 key='Replace_C101')],
    [sg.Checkbox('Replace Class 156 sets', default=config.getboolean('defaults', 'replace_c156'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Oovee Class 156s with AP Class 156', key='Replace_C156')],
    [sg.Checkbox('Replace Class 158 sets', default=config.getboolean('defaults', 'replace_c158'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of North Wales Coast / Settle Carlisle / Fife Circle Class 158s '
                         'with AP enhanced versions (Cummins, Perkins)',
                 key='Replace_C158')],
    [sg.Button('Replace!'), sg.Button('Settings'), sg.Button('About'), sg.Button('Exit')],
]

# Set the layout of the application window
layout = [
    [
        sg.Column(left_column),
        sg.VSeperator(),
        sg.Column(right_column),
    ]
]


def dcsv_get_num(this_dcsv, this_rv, this_re):
    # Try to retrieve the closest match for the loco number from the AP vehicle number database
    try:
        dcsv_tree = ET.parse(this_dcsv)
    except FileNotFoundError:
        sg.popup('AP vehicle number database ' + str(Path(this_dcsv)) + ' not found.',
                 'Check you have all the required AP products installed, and that you have clicked "Settings" in this '
                 'program and set the location of your RailWorks folder correctly.',
                 'This program will now quit.', title='Error')
        sys.exit('Fatal Error: AP vehicle number database ' + str(Path(this_dcsv)) + ' not found.')
    except ET.ParseError:
        sg.popup('AP vehicle number database ' + str(Path(this_dcsv)) + ' was found but could not be parsed.',
                 'This program will now quit.', title='Error')
        sys.exit(
            'Fatal Error: AP vehicle number database ' + str(Path(this_dcsv)) + ' was found but could not be parsed.')
    diff = 0
    last_nm = this_rv
    ithis_rv = int(this_rv)
    s = set(rv_list)
    root = dcsv_tree.getroot()
    for vnum in root.findall("./CSVItem/cCSVItem/Name"):
        # Iterate through the list of TOPS numbers until we find a number we haven't previously used which is an exact
        # match or the closest match for the number we're looking for
        nm = re.search(this_re, vnum.text)
        if nm:
            curr_nm = nm.group(1) + nm.group(2)
            dcsv_nm = int(nm.group(1))
            if curr_nm in s:
                # This number is already in use in another swapped loco - move on and try the next one
                continue
            if ithis_rv > dcsv_nm:
                # This number is still less than the number we're looking for - but remember how close it is and try the
                # next one to see if it is a match or is even further away than this number.
                diff = ithis_rv - dcsv_nm
            elif ithis_rv == dcsv_nm:
                # A matching number is available and has been found - use it.
                return curr_nm
            elif ithis_rv < dcsv_nm:
                # We have overshot the number we are looking for - but if this number is even further from the number
                # we're looking for than the last one we looked at, use the last one.
                if dcsv_nm - ithis_rv > diff:
                    return last_nm
                else:
                    # We've checked and even though we have overshot the number we are looking for, this number is
                    # closer to the number we're looking for than the last one we looked at - so use it.
                    return curr_nm
            last_nm = curr_nm
    # We didn't find a number to use. We must have reached the end of the available numbers so we will have to use the
    # last available number we found.
    return last_nm


def csv_get_blue47num(front, this_rv):
    ithis_rv = int(this_rv)
    last_loco = vp_blue_47_db[front][0]
    diff = 0
    s = set(rv_list)
    for loco in vp_blue_47_db[front]:
        # Iterate through the list of TOPS numbers until we find a number we haven't previously used which is an exact
        # match or the closest match for the number we're looking for, within the same subclass
        curr_nm = loco[0]
        dcsv_nm = int(loco[2])
        if curr_nm in s:
            # This number is already in use in another swapped loco - move on and try the next one
            continue
        if ithis_rv > dcsv_nm:
            # This number is still less than the number we're looking for - but remember how close it is and try the
            # next one to see if it is a match or is even further away than this number.
            diff = ithis_rv - dcsv_nm
        elif ithis_rv == dcsv_nm:
            # A matching number is available and has been found - use it.
            return loco
        elif ithis_rv < dcsv_nm:
            # We have overshot the number we are looking for - but if this number is even further from the number we're
            # looking for than the last one we looked at, use the last one.
            # Also, check to see if we have crossed the boundary into another subclass and if we have, remember the last
            # available number we found (from the previous subclass) and use that instead.
            if (dcsv_nm - ithis_rv > diff) or \
                    (ithis_rv < 47301 and dcsv_nm >= 47301) or \
                    (ithis_rv < 47401 and dcsv_nm >= 47401) or \
                    (ithis_rv < 47701 and dcsv_nm >= 47701):
                return last_loco
            else:
                # We've checked and this number is within the same subclass and even though we have overshot the number
                # we are looking for, this number is closer to the number we're looking for than the last one we
                # looked at - so use it.
                return loco
        last_loco = loco
    # We didn't find a number to use. We must have reached the end of the available numbers so we will have to use the
    # last available number we found.
    return last_loco


def dcsv_gethstloco(this_dcsv, this_rv):
    # Try to retrieve the closest match for the HST power car number from the AP vehicle number database
    try:
        dcsv_tree = ET.parse(this_dcsv)
    except FileNotFoundError:
        sg.popup('AP vehicle number database ' + str(Path(this_dcsv)) + ' not found.',
                 'Check you have all the required AP products installed, and that you have clicked "Settings" in this '
                 'program and set the location of your RailWorks folder correctly.',
                 'This program will now quit.', title='Error')
        sys.exit('Fatal Error: AP vehicle number database ' + str(Path(this_dcsv)) + ' not found.')
    except ET.ParseError:
        sg.popup('AP vehicle number database ' + str(Path(this_dcsv)) + ' was found but could not be parsed.',
                 'This program will now quit.', title='Error')
        sys.exit(
            'Fatal Error: AP vehicle number database ' + str(Path(this_dcsv)) + ' was found but could not be parsed.')
    diff = 0
    last_nm = this_rv
    irv = re.search('(43[0-9]{3})', this_rv)
    ithis_rv = int(irv.group(1))
    s = set(rv_list)
    root = dcsv_tree.getroot()
    for vnum in root.findall("./CSVItem/cCSVItem/Name"):
        # Iterate through the list of TOPS numbers until we find a number we haven't previously used which is an exact
        # match or the closest match for the number we're looking for
        nm = re.search('(.?)(43[0-9]{3})(.*)', vnum.text)
        if nm:
            curr_nm = nm.group(1) + nm.group(2) + nm.group(3)
            dcsv_nm = int(nm.group(2))
            if curr_nm in s:
                # This number is already in use in another swapped power car - move on and try the next one
                continue
            if ithis_rv > dcsv_nm:
                # This number is still less than the number we're looking for - but remember how close it is and try the
                # next one to see if it is a match or is even further away than this number.
                diff = ithis_rv - dcsv_nm
            elif ithis_rv == dcsv_nm:
                # A matching number is available and has been found - use it.
                return curr_nm
            elif ithis_rv < dcsv_nm:
                # We have overshot the number we are looking for - but if this number is even further from the number
                # we're looking for than the last one we looked at, use the last one.
                if dcsv_nm - ithis_rv > diff:
                    return last_nm
                else:
                    # We've checked and even though we have overshot the number we are looking for, this number is
                    # closer to the number we're looking for than the last one we looked at - so use it.
                    return curr_nm
            last_nm = curr_nm
    # We didn't find a number to use. We must have reached the end of the available numbers so we will have to use the
    # last available number we found.
    return last_nm


def add_ploughs(this_rv):
    # set full snow ploughs on AP locos that can have them
    this_rv = this_rv.replace(';plough=none', '')
    this_rv = this_rv.replace(';plough=outer', '')
    this_rv = this_rv.replace(';plough=full', '')
    this_rv = this_rv + ';plough=full'
    return this_rv


def add_retb(this_rv):
    # show the RETB equipment in Class 37 cab
    this_rv = this_rv.replace(';datacord=retb', '')
    this_rv = this_rv + ';datacord=retb'
    return this_rv


def get_coal21t_db(this_wagon):
    # list the available vehicle numbers for the Fastline Simulation 21T coal wagon
    if this_wagon == 'HTO 21t Hoppers - Dia 141':
        return HTO_141_numbers
    if this_wagon == 'HTO 21t Hoppers - Dia 143':
        return HTO_143_numbers
    if this_wagon == 'HTO 21t Hoppers - Dia 146':
        return HTO_146_numbers
    if this_wagon == 'HTO 21t Hoppers - Rebodied':
        return HTO_rebodied_numbers
    if this_wagon == 'HTV 21t Hoppers - Dia 146':
        return HTV_146_numbers
    if this_wagon == 'HTV 21t Hoppers - Rebodied':
        return HTV_rebodied_numbers
    return False


def dcsv_21t_hopper_number(this_rv, this_rv_list):
    # return a wagon number in Fastline format for the various 21t coal hoppers
    rv_digits = int(re.sub('[^0-9]]', "", this_rv))
    for i in range(0, len(this_rv_list)):
        if 'B' + str(rv_digits) == this_rv_list[i]:
            return this_rv_list[i]
    # No exact match was found - use modulus operator to select one
    return this_rv_list[rv_digits % len(this_rv_list)]


def cl50char_to_num(this_rv):
    # This converts the Meshtools single character identifier to the AP loco number
    i = 50000
    a = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
         'N', 'M', 'O', 'P', 'Q', 'R', 'S', 'T', 'W', 'U', 'X', 'Y', 'Z', '!', '£', '$', '%', '-', '_', '=', '+', '[',
         '{', ']', '}', '#', '~', '@']
    for loco in a:
        i = i + 1
        if loco == this_rv:
            return str(i)
    # If no match found you get 50043 'Eagle'
    return 50043


def cl56rsc_to_apsecdep_or_blanksecdep(this_rv):
    # This converts the RSC Class 56 RF number, sector and depot to AP numbering scheme
    # sector and/or plaque blank where no AP equivalent exists
    this_sector = this_rv[0:1]
    this_depot = this_rv[1:2]
    sectors = {'a': 'd', 'b': 'c', 'e': 'b', 'f': 'a'}
    if this_sector in sectors:
        this_sector = sectors[this_sector]
    else:
        this_sector = '*'
    depots = {'G': 'C', 'V': 'D', 'X': 'A'}
    if this_depot in depots:
        this_depot = depots[this_depot]
    else:
        this_depot = '*'
    return this_sector + this_depot + this_rv[2:7]


def set_weathering(this_weather_variant, this_vehicle):
    if this_weather_variant == 2:
        weather = 'W' + str(random.randint(1, 2))
        return this_vehicle[5].replace('W2', weather), this_vehicle[6].replace('W2', weather)
    elif this_weather_variant == 3:
        weather = 'W' + str(random.randint(1, 3))
        return this_vehicle[5].replace('W1', weather), this_vehicle[6].replace('W1', weather)
    return False


def haa_replace(provider, product, blueprint, name, number, loaded):
    # Replace HAA wagons
    for i in range(0, len(vehicle_db['HAA'])):
        this_vehicle = vehicle_db['HAA'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = 'AP'
                    product.text = 'HAAWagonPack01'
                    # Replace a loaded wagon
                    if 'eTrue' in loaded.text:
                        idx = random.randrange(0, len(haa_l_wagons))
                        # Select at random one of the wagons in the list of HAA loaded wagons to swap in
                        blueprint.text = haa_l_wagons[idx][0]
                        name.text = haa_l_wagons[idx][1]
                    # Replace an empty wagon
                    if 'eFalse' in loaded.text:
                        idx = random.randrange(0, len(haa_e_wagons))
                        # Select at random one of the wagons in the list of HAA empty wagons to swap in
                        blueprint.text = haa_e_wagons[idx][0]
                        name.text = haa_e_wagons[idx][1]
                    # Now extract the vehicle number
                    rv_list.append(number.text)
                    return True
    return False


def fsafta_replace(provider, product, blueprint, name, number, loaded):
    if bool(fsa_replace(provider, product, blueprint, name, number, loaded)):
        return True
    if bool(fta_replace(provider, product, blueprint, name, number, loaded)):
        return True
    return False


def fsa_replace(provider, product, blueprint, name, number, loaded):
    for i in range(0, len(vehicle_db['FSA'])):
        this_vehicle = vehicle_db['FSA'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    if 'eFalse' in loaded.text:
                        # Wagon is unloaded
                        number.text = dcsv_get_num(
                            Path(railworks_path, 'Assets/AP/FSAWagonPack/RailVehicles/Freight/FL/FSA.dcsv'),
                            number.text,
                            '([0-9]{6})(.*)')
                        # Change the blueprint and name to the unloaded wagon
                        blueprint.text = re.sub('FSA[a-zA-Z0-9_]*.xml', 'FSA.xml', this_vehicle[5], flags=re.IGNORECASE)
                        name.text = re.sub('AP.FSA.([a-zA-Z]*).*', r'AP FSA \1', this_vehicle[6], flags=re.IGNORECASE)
                    else:
                        number.text = dcsv_get_num(
                            Path(railworks_path, 'Assets/AP/FSAWagonPack', this_vehicle[7]), number.text,
                            '([0-9]{6})(.*)')
                    rv_pairs.append([rv_orig, number.text])
                    rv_list.append(number.text)
                    return True
    return False


def fta_replace(provider, product, blueprint, name, number, loaded):
    for i in range(0, len(vehicle_db['FTA'])):
        this_vehicle = vehicle_db['FTA'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    if 'eFalse' in loaded.text:
                        # Wagon is unloaded
                        number.text = dcsv_get_num(
                            Path(railworks_path, 'Assets/AP/FSAWagonPack/RailVehicles/Freight/FL/FTA.dcsv'),
                            number.text,
                            '([0-9]{6})(.*)')
                        # Change the blueprint and name to the unloaded wagon
                        blueprint.text = re.sub('FTA[a-zA-Z0-9_]*.xml', 'FTA.xml', this_vehicle[5], flags=re.IGNORECASE)
                        name.text = re.sub('AP.FTA.([a-zA-Z]*).*', r'AP FTA \1', this_vehicle[6], flags=re.IGNORECASE)
                    else:
                        number.text = dcsv_get_num(
                            Path(railworks_path, 'Assets/AP/FSAWagonPack', this_vehicle[7]), number.text,
                            '([0-9]{6})(.*)')
                    rv_pairs.append([rv_orig, number.text])
                    rv_list.append(number.text)
                    return True
    return False


def mk1_replace(provider, product, blueprint, name, number):
    # Replace any Mk1s - loop through the VehicleDB['Mk1'] array of coaches to search for
    for i in range(0, len(vehicle_db['Mk1'])):
        this_vehicle = vehicle_db['Mk1'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    # Now extract the region code (if there is one) and the coach number
                    nm = re.search('([a-zA-Z]{0,2})([0-9]{4,5})', number.text)
                    if nm:
                        rv_orig = number.text
                        region = nm.group(1).upper()
                        num = nm.group(2)
                        # Express the region prefix (or lack of) in a manner compatible with AP numbering scheme
                        if region == 'E' or region == 'S' or region == 'W' or region == 'SC':
                            ap_suffix = ';R=' + region
                        elif len(region) < 1:
                            ap_suffix = ';R=Z'
                        else:
                            ap_suffix = ''
                        num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                            num, '([0-9]{4,5})(.*)')
                        if ' (Newspapers)' in name.text:
                            # Add the AP coach number suffix to display the Newspapers branding on BG coaches
                            ap_suffix = ap_suffix + ";L=6"
                        elif ' (Parcels)' in name.text:
                            # Add the AP coach number suffix to display the Parcels branding on BG coaches
                            ap_suffix = ap_suffix + ";L=3"
                        elif ' (ScotRail)' in name.text:
                            # Add the AP coach number suffix to display the ScotRail branding on BR Blue/Grey coaches
                            ap_suffix = ";R=SC;L=5"
                        elif ' (Swallow)' in name.text:
                            # Add the AP coach number suffix to display the Swallow brand on InterCity coaches
                            ap_suffix = ap_suffix + ";L=2"
                        elif 'BR Blue/Grey (NSE)' in name.text:
                            # Add the AP coach number suffix to display the NSE branding on BR Blue/Grey coaches
                            ap_suffix = ";L=2"
                        elif ' (unbranded)' in name.text:
                            # Add the AP coach number suffix to remove logos
                            ap_suffix = ";L=0"
                        rv_num = num + ap_suffix
                        number.text = rv_num
                        rv_pairs.append([rv_orig, number.text])
                        rv_list.append(number.text)
                        # Following line sets AP coach Number
                    return True
    return False


def mk2ac_replace(provider, product, blueprint, name, number):
    # Replace any Mk2a/b/cs - loop through the VehicleDB['Mk2ac'] array of coaches to search for
    for i in range(0, len(vehicle_db['Mk2ac'])):
        this_vehicle = vehicle_db['Mk2ac'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    # Now extract the region code (if there is one) and the coach number
                    nm = re.search('([a-zA-Z]{0,2})([0-9]{4,5})', number.text)
                    if nm:
                        rv_orig = number.text
                        region = nm.group(1).upper()
                        num = nm.group(2)
                        ap_suffix = ''
                        # Express the region prefix (or lack of) in a manner compatible with AP numbering scheme
                        if region == 'E' or region == 'S' or region == 'W' or region == 'SC':
                            ap_suffix = ";R=" + region
                        elif len(region) < 1:
                            ap_suffix = ";R=Z"
                        if 'BR Blue/Grey NSE' in name.text:
                            # Add the AP coach number suffix to display the BR Blue/Grey NSE branding
                            ap_suffix = ap_suffix + ";L=2"
                        elif 'VintageTrains' in name.text:
                            # Add the AP coach number suffix to remove branding as per Vintage Trains
                            ap_suffix = ap_suffix + ";L=0"
                        elif 'BR Blue/Grey ScotRail' in name.text:
                            # Add the AP coach number suffix to display the BR Blue/Grey ScotRail branding
                            nm = re.search('R=[^Z]', ap_suffix)
                            if nm:
                                # If the original has a non-Scottish region, change it to Sc
                                ap_suffix = ";R=SC;L=3"
                            else:
                                # If the original has no region letter, leave it with no region
                                ap_suffix = ";R=Z;L=3"
                        rv_num = num + ap_suffix
                        rv_pairs.append([rv_orig, rv_num])
                        rv_list.append(rv_num)
                        # Following line sets AP coach number
                        number.text = rv_num
                    return True
    return False


def mk2df_replace(provider, product, blueprint, name, number):
    # Replace any Mk2d/e/fs - loop through the VehicleDB['Mk2df'] array of coaches to search for
    for i in range(0, len(vehicle_db['Mk2df'])):
        this_vehicle = vehicle_db['Mk2df'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    # Now extract the region code (if there is one) and the coach number
                    nm = re.search('([0-9]{4,5})', number.text)
                    if nm:
                        num = nm.group(1)
                        ap_suffix = ";R=Z"
                        rv_num = num + ap_suffix
                        rv_pairs.append([rv_orig, rv_num])
                        rv_list.append(rv_num)
                        # Following line sets AP coach number
                        number.text = rv_num
                    return True
    return False


def vda_replace(provider, product, blueprint, name, number, loaded):
    # Replace VDA wagons
    if 'JL' in provider.text:
        if 'WHL' in product.text:
            bp = re.search(re.escape(r'RailVehicles\Freight\VDA\VDA.xml'), blueprint.text, flags=re.IGNORECASE)
            if bp:
                provider.text = 'FastlineSimulation'
                rv_orig = number.text
                if 'eTrue' in loaded.text:
                    # Replace a loaded wagon
                    idx = random.randrange(0, len(vda_l_wagons))
                    product.text = vda_l_wagons[idx][1]
                    blueprint.text = vda_l_wagons[idx][2]
                    name.text = vda_l_wagons[idx][3]
                else:
                    # Replace an empty wagon
                    idx = random.randrange(0, len(vda_e_wagons))
                    product.text = vda_e_wagons[idx][1]
                    blueprint.text = vda_e_wagons[idx][2]
                    name.text = vda_e_wagons[idx][3]
                # Now process the vehicle number
                rv_num = rv_orig + "#####"
                rv_list.append(rv_num)
                rv_pairs.append([rv_orig, rv_num])
                # Set Fastline wagon number
                number.text = rv_num
                return True
    return False


def coal21_t_hto_replace(provider, product, blueprint, name, number, loaded):
    for i in range(0, len(vehicle_db['HTO'])):
        this_vehicle = vehicle_db['HTO'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = 'FastlineSimulation'
                    rv_orig = number.text
                    if 'eTrue' in loaded.text:
                        # Replace a loaded wagon
                        idx = random.randrange(0, len(hto_l_wagons))
                        product.text = hto_l_wagons[idx][1]
                        blueprint.text = hto_l_wagons[idx][2]
                        name.text = hto_l_wagons[idx][3]
                        # Now process the vehicle number
                        rv_num = dcsv_21t_hopper_number(rv_orig, get_coal21t_db(hto_l_wagons[idx][1]))
                        rv_list.append(rv_num)
                        rv_pairs.append([rv_orig, rv_num])
                        # Set Fastline wagon number
                        number.text = str(rv_num)
                        return True
                    else:
                        # Replace an empty wagon
                        idx = random.randrange(0, len(hto_e_wagons))
                        product.text = hto_e_wagons[idx][1]
                        blueprint.text = hto_e_wagons[idx][2]
                        name.text = hto_e_wagons[idx][3]
                        # Now process the vehicle number
                        rv_num = dcsv_21t_hopper_number(rv_orig, get_coal21t_db(hto_e_wagons[idx][1]))
                        rv_list.append(rv_num)
                        rv_pairs.append([rv_orig, rv_num])
                        # Set Fastline wagon number
                        number.text = str(rv_num)
                    return True
    return False


def coal21_t_htv_replace(provider, product, blueprint, name, number, loaded):
    for i in range(0, len(vehicle_db['HTV'])):
        # Replace fitted wagons
        this_vehicle = vehicle_db['HTV'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = 'FastlineSimulation'
                    rv_orig = number.text
                    if 'eTrue' in loaded.text:
                        # Replace a loaded wagon
                        idx = random.randrange(0, len(htv_l_wagons))
                        product.text = htv_l_wagons[idx][1]
                        blueprint.text = htv_l_wagons[idx][2]
                        name.text = htv_l_wagons[idx][3]
                        # Now process the vehicle number
                        rv_num = dcsv_21t_hopper_number(rv_orig, get_coal21t_db(htv_l_wagons[idx][1]))
                        rv_list.append(rv_num)
                        rv_pairs.append([rv_orig, rv_num])
                        # Set Fastline wagon number
                        number.text = str(rv_num)
                        return True
                    else:
                        # Replace an empty wagon
                        idx = random.randrange(0, len(htv_e_wagons))
                        product.text = htv_e_wagons[idx][1]
                        blueprint.text = htv_e_wagons[idx][2]
                        name.text = htv_e_wagons[idx][3]
                        # Now process the vehicle number
                        rv_num = dcsv_21t_hopper_number(rv_orig, get_coal21t_db(htv_e_wagons[idx][1]))
                        rv_list.append(rv_num)
                        rv_pairs.append([rv_orig, rv_num])
                        # Set Fastline wagon number
                        number.text = str(rv_num)
                    return True
    return False


def ihh_replace(provider, product, blueprint, name, number):
    if bool(ihh_bonus_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c20_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c25_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c26_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c27_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c40_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c45_replace(provider, product, blueprint, name, number)):
        return True
    return False


def ihh_bonus_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['IHH_Bonus'])):
        this_vehicle = vehicle_db['IHH_Bonus'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text
                    guv = re.search('guv', blueprint.text, flags=re.IGNORECASE)
                    cao = re.search('20t', blueprint.text, flags=re.IGNORECASE)
                    c47 = re.search('brush_4_bue', blueprint.text, flags=re.IGNORECASE)
                    if guv:
                        if not bool(re.match('[A-Z][0-9]{5}', number.text)):
                            # If the number is not in expected format, choose a random one.
                            number.text = 'M' + str(random.randint(86078, 86984))
                    elif cao:
                        # As the IHH number format for 20t brake vans is not known, choose a random number.
                        number.text = '####B' + str(random.randint(953676, 954520)) + '#'
                    elif c47:
                        # Initialise a random Class 47/0 number in case no valid number found
                        rv_num = str(random.randint(47001, 47298))
                        # Try to extract loco number from IHH number string
                        nm_tops = re.search('^47#([0-9]{3}).*', number.text)
                        nm_pretops = re.search('^D#([0-9]{4}).*', number.text)
                        if nm_tops:
                            rv_num = str(47000 + int(nm_tops.group(1)))
                        elif nm_pretops:
                            # It's a pre-tops number - select a 47/0 TOPS number instead
                            rv_num = str(47001 + ((int(nm_pretops.group(1)) - 1) % 298))
                        # look up the TOPS number and retrieve details for VP blueprints and number
                        loco = csv_get_blue47num(this_vehicle[3], rv_num)
                        this_vehicle[3] = 'Kuju'
                        this_vehicle[4] = 'RailSimulator'
                        this_vehicle[5] = loco[4]
                        this_vehicle[6] = loco[3]
                        number.text = loco[0]
                    else:
                        return False
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def ihh_c20_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class 20' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class20'])):
                this_vehicle = vehicle_db['IHH_Class20'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text
                    rv_num = str(random.randint(20001, 20126))
                    nm = re.search('^.20#([0-9]{3})', number.text)
                    if nm:
                        if int(nm.group(1)) < 127:
                            rv_num = str(20000 + int(nm.group(1)))
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = rv_num
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def ihh_c25_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_25' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class25'])):
                this_vehicle = vehicle_db['IHH_Class25'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_num = '251040000'
                    rv_orig = number.text
                    nm = re.search('^(25[0-9]{3})(....).*', number.text)
                    if nm:
                        rv_num = nm.group(1)
                        headcode = '0000'
                        hc_search = re.search('([0-9][A-Z][0-9]{2})', nm.group(2))
                        if hc_search:
                            headcode = nm.group(2)
                        elif this_vehicle[3] == 'RSderek':
                            headcode = '@##@'
                        rv_num = rv_num + headcode
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = rv_num
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def ihh_c26_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_26' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class26'])):
                this_vehicle = vehicle_db['IHH_Class26'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_num = '26024'
                    rv_orig = number.text
                    nm = re.search('^(26[0-9]{3}).*', number.text)
                    if nm:
                        rv_num = nm.group(1)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = rv_num
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def ihh_c27_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_27' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class27'])):
                this_vehicle = vehicle_db['IHH_Class27'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_num = '27024'
                    rv_orig = number.text
                    nm = re.search('^(27[0-9]{3}).*', number.text)
                    if nm:
                        rv_num = nm.group(1)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = rv_num
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def ihh_c40_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_40' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class40'])):
                this_vehicle = vehicle_db['IHH_Class40'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_num = number.text
                    rv_orig = number.text
                    if bool(re.search('disc_blue|late_blue', this_vehicle[2], flags=re.IGNORECASE)):
                        tops_disc = re.search('^(40[0-9]{3})(....).*', number.text)
                        if tops_disc:
                            rv_tops = '1111' + tops_disc.group(1)
                            ap_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                                rv_tops,
                                '([0-9]{9})(.*)')
                            rv_num = ap_num[0:9] + '2222'
                    else:
                        tops_headcode = re.search('^(40[0-9]{3})(....).*', number.text)
                        if tops_headcode:
                            rv_tops = '11111' + tops_headcode.group(1)
                            ap_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                                rv_tops, '([0-9]{10})(.*)')
                            hc_search = re.search('([0-9][A-Z][0-9]{2})', tops_headcode.group(2))
                            if hc_search:
                                rv_num = '110' + ap_num[3:10] + hc_search.group(0)
                            else:
                                rv_num = ap_num
                    # Set AP Class 40 number
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def ihh_c45_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_45' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class45'])):
                this_vehicle = vehicle_db['IHH_Class45'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_num = number.text
                    rv_orig = number.text
                    nm = re.search('^(45|46)#([0-9]{3}).*', number.text)
                    if nm:
                        rv_num = nm.group(1) + nm.group(2)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = rv_num
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def hst_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['HST_set'])):
        this_vehicle = vehicle_db['HST_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    # Now extract the vehicle number
                    if 'Class43' in this_vehicle[5]:
                        nm = re.search('(.?43[0-9]{3}.*)', number.text)
                        if nm:
                            rv_num = dcsv_gethstloco(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                                number.text)
                            number.text = str(rv_num)
                            rv_list.append(number.text)
                            rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c31_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Class31'])):
        this_vehicle = vehicle_db['Class31'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    if 'W2' in this_vehicle[5]:
                        (w_blueprint, w_name) = set_weathering(2, this_vehicle)
                    else:
                        (w_blueprint, w_name) = set_weathering(3, this_vehicle)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = w_blueprint
                    name.text = w_name
                    rv_orig = number.text
                    nm = re.search('[^3]*(31[0-9]{3}).*', number.text)
                    if nm:
                        rv_found = nm.group(1)
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]), rv_found,
                            '([0-9]{5})(.*)')
                        # Set number
                        number.text = str(rv_num)
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c37_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Class37'])):
        this_vehicle = vehicle_db['Class37'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    (w_blueprint, w_name) = set_weathering(3, this_vehicle)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = w_blueprint
                    name.text = w_name
                    rv_num = rv_tops = rv_orig = number.text
                    # Check if the loco has a pre-tops number
                    pretops = re.search('D([0-9]{4})([0-9][a-zA-Z][0-9]{2})', number.text)
                    if pretops:
                        rv_dnum = pretops.group(1)
                        headcode = pretops.group(2)
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]), rv_dnum,
                            '([0-9]{5})(.*)')
                        rv_num = rv_num.replace('____', headcode)
                    # Check if the loco has a tops number
                    tops = re.search('(37[0-9]{3})(.*)', number.text)
                    if tops:
                        rv_tops = tops.group(1)
                        rv_num = this_vehicle[7]
                        if 'dcsv' in rv_num:
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                                rv_tops,
                                '([0-9]{5})(.*)')
                    if '_wp' in this_vehicle[2]:
                        add_ploughs(rv_num)
                    if this_vehicle[1] == 'WHL' or this_vehicle[1] == 'FortWilliamMallaig':
                        if 'Large' in this_vehicle[2]:
                            # Look for a loco with the 'Westie' logo for the WHL LL replacements
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                                rv_tops,
                                '(37[0-9]{3})(.*L=1.*)')
                        # Add ploughs and RETB to West Highland locos
                        add_retb(rv_num)
                        add_ploughs(rv_num)
                        if 'Default' in this_vehicle[2]:
                            # Black headcode box
                            rv_num = rv_num + ';no1front=bch;no2front=bch'
                    # Set AP Class 37 number
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c40_replace(provider, product, blueprint, name, number):
    # Replace DT Class 40
    if 'DT' in provider.text:
        if 'DT_class40' in product.text:
            for i in range(0, len(vehicle_db['Class40'])):
                this_vehicle = vehicle_db['Class40'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_num = number.text
                    rv_orig = number.text
                    # Check if the loco has a pre-tops number
                    pretops_disc = re.search('^([0-9])([0-9]{3})$', number.text)
                    if pretops_disc:
                        rv_dnum = '0' + pretops_disc.group(2)
                        headcode = ap40headcodes_62_69[pretops_disc.group(1)]
                        ap_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                            rv_dnum,
                            '([0-9]{4})(.*)')
                        hy = re.search('halfyellow', blueprint.text, flags=re.IGNORECASE)
                        if hy:
                            # Number the loco as a Half Yellow front class 40
                            rv_num = '1' + ap_num[1:4] + headcode
                        else:
                            # Number the loco as a Full Green front class 40
                            rv_num = '0' + ap_num[1:4] + headcode
                    pretops_headcode = re.search('^([0-9][a-z][0-9]{2})([0-9]{3})$', number.text)
                    if pretops_headcode:
                        rv_dnum = '0' + pretops_headcode.group(2)
                        headcode = pretops_headcode.group(1).upper()
                        ap_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                            rv_dnum,
                            '([0-9]{4})(.*)')
                        hy = re.search('halfyellow', blueprint.text, flags=re.IGNORECASE)
                        if hy:
                            # Number the loco as a Half Yellow front class 40
                            rv_num = '1' + ap_num[1:4] + headcode
                        else:
                            # Number the loco as a Full Green front class 40
                            rv_num = '0' + ap_num[1:4] + headcode
                    # Check if the loco has a tops number
                    tops_domino = re.search('^(40[0-9]{3})$', number.text)
                    if tops_domino:
                        rv_tops = '11111' + tops_domino.group(1)
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                            rv_tops,
                            '([0-9]{10})(.*)')
                    tops_disc = re.search('^([0-9])(40[0-9]{3})$', number.text)
                    if tops_disc:
                        rv_tops = '1111' + tops_disc.group(2)
                        headcode = ap40headcodes_69_77[tops_disc.group(1)]
                        ap_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                            rv_tops,
                            '([0-9]{9})(.*)')
                        rv_num = ap_num[0:9] + headcode
                    # Set AP Class 40 number
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c47_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Class47BRBlue'])):
        this_vehicle = vehicle_db['Class47BRBlue'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text
                    nm = re.search('^(47[0-9]{3})', number.text)
                    if nm:
                        loco = csv_get_blue47num(this_vehicle[3], nm.group(1))
                        provider.text = 'Kuju'
                        product.text = 'RailSimulator'
                        blueprint.text = loco[4]
                        name.text = loco[3]
                        number.text = loco[0]
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                        return True
                    else:
                        return False
    return False


def c50_replace(provider, product, blueprint, name, number):
    # Replace MT Class 50
    if 'MichaelWhiteley' in provider.text:
        if 'Class 50' in product.text:
            for i in range(0, len(vehicle_db['Class50'])):
                this_vehicle = vehicle_db['Class50'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                pretops = re.match('([0-9]{3})([0-9][a-zA-Z][0-9]{2})', number.text)
                if bp:
                    (w_blueprint, w_name) = set_weathering(3, this_vehicle)
                    rv_orig = number.text
                    if i == 0:
                        # This is the GWR loco, 50007 - note that it only has two weathered variants, W1 and W2
                        (w_blueprint, w_name) = set_weathering(2, this_vehicle)
                        rv_num = '50007'
                        rv_list.append(rv_num)
                    elif len(number.text) == 1:
                        # This is one of the BR/NSE TOPS liveries with single character vehicle number to translate
                        rv_num = cl50char_to_num(rv_orig)
                        rv_list.append(rv_num)
                        rv_pairs.append([rv_orig, rv_num])
                    elif len(pretops.group(0)) == 6:
                        # This is one of the headcode box variants - pre-TOPS BR Green or Blue
                        rv_num = 'D' + pretops.group(1) + ';L=1;HC1=' + pretops.group(2) + ';HC2=' + pretops.group(2)
                        rv_list.append(rv_num)
                        rv_pairs.append([rv_orig, rv_num])
                    else:
                        rv_num = rv_orig
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = w_blueprint
                    name.text = w_name
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c56_replace(provider, product, blueprint, name, number):
    # Replace RSC Class 56 with AP Enhanced version
    if 'RSC' in provider.text:
        if 'Class56Pack01' in product.text:
            for i in range(0, len(vehicle_db['Class56'])):
                this_vehicle = vehicle_db['Class56'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_num = cl56rsc_to_apsecdep_or_blanksecdep(number.text)
                    rv_orig = number.text
                    if config.get('defaults', 'c56_rf') == c56_opts[1]:
                        # Skip swapping in AP loco unless it has both the sectors logo and depot plaque
                        # of the loco it is to replace
                        if rv_num[0:1] == '*':
                            return False
                        if rv_num[1:2] == '*':
                            return False
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c66_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Class66'])):
        this_vehicle = vehicle_db['Class66'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    nm = re.search('(66[0-9]{3}).*', number.text)
                    if nm:
                        rv_found = nm.group(1)
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]), rv_found,
                            '([0-9]{5})(.*)')
                        # Set number
                        number.text = str(rv_num)
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c67_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Class67'])):
        this_vehicle = vehicle_db['Class67'][i]
        if this_vehicle[0] in provider.text and this_vehicle[1] in product.text:
            bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
            if bp:
                (w_blueprint, w_name) = set_weathering(3, this_vehicle)
                provider.text = this_vehicle[3]
                product.text = this_vehicle[4]
                blueprint.text = w_blueprint
                name.text = w_name
                rv_orig = number.text
                nm = re.search('(67[0-9]{3}).*', number.text)
                if nm:
                    rv_found = number.text
                    rv_num = this_vehicle[7]
                    if 'dcsv' in rv_num:
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]), rv_found,
                            '([0-9]{5})(.*)')
                    # Set number
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                return True
    return False


def c68_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Class68'])):
        this_vehicle = vehicle_db['Class68'][i]
        if this_vehicle[0] in provider.text and this_vehicle[1] in product.text:
            bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
            if bp:
                provider.text = this_vehicle[3]
                product.text = this_vehicle[4]
                blueprint.text = this_vehicle[5]
                name.text = this_vehicle[6]
                rv_orig = number.text
                nm = re.search('(68[0-9]{3}).*', number.text)
                if nm:
                    rv_found = nm.group(1)
                    rv_num = dcsv_get_num(
                        Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]), rv_found,
                        '([0-9]{5})(.*)')
                    # Set number
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                return True
    return False


def c86_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Class86'])):
        this_vehicle = vehicle_db['Class86'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    (w_blueprint, w_name) = set_weathering(3, this_vehicle)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = w_blueprint
                    name.text = w_name
                    rv_orig = number.text
                    nm = re.match('(86[0-9]{3}).*', number.text)
                    if nm:
                        # Loco to be replaced is TOPS numbered class 86 with no headcode box
                        rv_found = nm.group(1)
                        rv_num = this_vehicle[7]
                        if 'dcsv' in rv_num:
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4], this_vehicle[7]),
                                rv_found,
                                '([0-9]{5})(.*)')
                        # Set number
                        number.text = str(rv_num)
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                        return True
                    nm = re.search('([0-9][a-zA-Z][0-9]{2})(86[0-9]{3})', number.text)
                    if nm:
                        # This has the number of a RSC Class 86 BR Blue with TOPS number and headcode box
                        # Replace with Vulcan Productions headcode loco if the user asked for it - otherwise
                        # do nothing or swap for the standard AP BR Blue 1 with no headcode
                        if config.get('defaults', 'c86_hc') == c86_opts[2]:
                            # User doesn't want this loco replaced
                            return True
                        elif config.get('defaults', 'c86_hc') == c86_opts[0]:
                            # User wants this local replaced with the Vulcan Productions Class 86 Early Liveries &
                            # Headcode Blinds loco from https://www.vulcanproductions.co.uk/electric.html Note there
                            # is no dead / low panto version
                            blueprint.text = vehicle_db['Class86'][0][5]
                            name.text = vehicle_db['Class86'][0][6]
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', vehicle_db['Class86'][0][3], vehicle_db['Class86'][0][4],
                                     vehicle_db['Class86'][0][7]), nm.group(2), '([0-9]{5})(.*)')
                            # Set the headcode into the VP number
                            rv_num = rv_num.replace('0O00', nm.group(1))
                            # Set number
                            number.text = str(rv_num)
                            rv_list.append(number.text)
                            rv_pairs.append([rv_orig, number.text])
                            return True
                        elif config.get('defaults', 'c86_hc') == c86_opts[1]:
                            # User wants this loco replaced with the AP BR Blue 1 loco (no headcode blinds)
                            if 'panto_low' in this_vehicle[2]:
                                (w_blueprint, w_name) = set_weathering(3, vehicle_db['Class86'][4])
                            else:
                                (w_blueprint, w_name) = set_weathering(3, vehicle_db['Class86'][6])
                            blueprint.text = w_blueprint
                            name.text = w_name
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', vehicle_db['Class86'][4][3], vehicle_db['Class86'][4][4],
                                     vehicle_db['Class86'][4][7]), nm.group(2), '([0-9]{5})(.*)')
                            # Set number
                            number.text = str(rv_num)
                            rv_list.append(number.text)
                            rv_pairs.append([rv_orig, number.text])
                            return True
                    return False
    return False


def c91_replace(provider, product, blueprint, name):
    for i in range(0, len(vehicle_db['Class91_set'])):
        this_vehicle = vehicle_db['Class91_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    return True
    return False


def c101_replace(provider, product, blueprint, name):
    for i in range(0, len(vehicle_db['DMU101_set'])):
        this_vehicle = vehicle_db['DMU101_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    return True
    return False


def c156_replace(provider, product, blueprint, name, number):
    if 'Oovee' in provider.text:
        if 'BRClass156Pack01' in product.text:
            for i in range(0, len(vehicle_db['DMU156_set'])):
                this_vehicle = vehicle_db['DMU156_set'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text
                    nm = re.search('(156[0-9]{3})', number.text)
                    if nm:
                        number.text = nm.group(1) + 'a' + this_vehicle[7]
                    else:
                        # Unit number of the Oovee 156 is not in standard 156xxx format - can't replace the vehicle
                        return False
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c158_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['DMU158_set'])):
        this_vehicle = vehicle_db['DMU158_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = rv_num = number.text
                    destination = 'a'
                    if provider.text == 'S9Bl':
                        nm = re.search('(....).....([0-9]{6})', number.text)
                        if nm:
                            if bool(re.search('Default', blueprint.text, flags=re.IGNORECASE)):
                                destination = c158_s9bl_rr[nm.group(1)]
                            elif bool(re.search('FGW', blueprint.text, flags=re.IGNORECASE)):
                                destination = c158_s9bl_fgw[nm.group(1)]
                            elif bool(re.search('NR', blueprint.text, flags=re.IGNORECASE)):
                                destination = c158_s9bl_nr[nm.group(1)]
                            elif bool(re.search('NTPE', blueprint.text, flags=re.IGNORECASE)):
                                destination = c158_s9bl_tpe[nm.group(1)]
                            elif bool(re.search('South|SWT', blueprint.text, flags=re.IGNORECASE)):
                                destination = c158_s9bl_swt[nm.group(1)]
                            rv_num = nm.group(2) + destination
                    else:
                        nm = re.search('(.)([0-9]{4}).*', number.text)
                        if nm:
                            if (provider.text == 'DTG' and product.text == 'Class158Pack01' and bool(
                                    re.search('Default', blueprint.text, flags=re.IGNORECASE))) or (
                                    provider.text == 'DTG' and product.text == 'NorthWalesCoast' and bool(
                                re.search('Default', blueprint.text, flags=re.IGNORECASE))):
                                # Arriva Trains Wales liveried stock
                                destination = c158_nwc[nm.group(1)]
                            elif provider.text == 'DTG' and product.text == 'FifeCircle' and bool(
                                    re.search('Default', blueprint.text, flags=re.IGNORECASE)):
                                # ScotRail saltire liveried stock
                                destination = c158_dtg_fc[nm.group(1)]
                            elif provider.text == 'RSC' and product.text == 'LiverpoolManchester' and bool(
                                    re.search('Default', blueprint.text, flags=re.IGNORECASE)):
                                # Regional Railways liveried stock
                                destination = c158_livman_rr[nm.group(1)]
                            rv_num = '15' + nm.group(2) + destination
                        if provider.text == 'RSC' and product.text == 'SettleCarlisle':
                            # Destination blank - Settle-Carlisle units don't support destination displays
                            rv_num = '158' + rv_orig[2:5] + 'a'
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def user_replace(provider, product, blueprint, name):
    for i in range(0, len(user_db['User'])):
        this_vehicle = user_db['User'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    return True
    return False


def parse_xml(xml_file):
    # Check we can open the file, parse it and find some rail vehicle consists in it before proceeding
    try:
        parser_tree = ET.parse(xml_file)
    except FileNotFoundError:
        sg.popup('Scenario file ' + str(Path(xml_file)) + ' not found.', 'Please try again.', title='Error')
        return False
    except ET.ParseError:
        sg.popup('The file you requested (' + str(Path(xml_file)) + ') could not be processed due to an XML parse '
                                                                    'error. Is it definitely a scenario file?',
                 'Please try again with another Scenario.bin or Scenario.xml '
                 'file.', title='Error')
        return False
    ET.register_namespace("d", "http://www.kuju.com/TnT/2003/Delta")
    root = parser_tree.getroot()
    consists = root.findall('./Record/cConsist')
    if len(consists) == 0:
        sg.popup('The file you requested (' + str(Path(xml_file)) + ') does not appear to contain any rail vehicle '
                                                                    'consists. Is it definitely a scenario file?',
                 'Please try again with another Scenario.bin or Scenario.xml '
                 'file.', title='Error')
        return False
    # Iterate through the consists - pop up a progress bar window
    progress_layout = [
        [sg.Text('Processing consists')],
        [sg.ProgressBar(1, orientation='h', key='progress', size=(25, 15))]
    ]
    progress_win = sg.Window('Processing...', progress_layout, disable_close=True).Finalize()
    progress_bar = progress_win.FindElement('progress')
    consist_nr = 0
    for citem in consists:
        service = citem.find('Driver/cDriver/ServiceName/Localisation-cUserLocalisedString/English')
        if service is None:
            service = 'Loose consist'
        else:
            service = service.text
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
                vehicle_replacer(provider, product, blueprint, name, number, loaded)
                vehicle_list.append(
                    [str(consist_nr), provider.text, product.text, blueprint.text, name.text, number.text, loaded.text,
                     service])
            consist_nr += 1
            progress_bar.UpdateBar(consist_nr, len(consists))
        for driver_inrvs in citem.findall('Driver/cDriver/InitialRV'):
            # Iterate through driver instructions and update changed vehicle numbers in the consist
            for drv in driver_inrvs.findall('e'):
                for rvp in rv_pairs:
                    if drv.text == rvp[0]:
                        drv.text = rvp[1]
    for citem in root.findall('./Record/cConsist'):
        # Now that the consist rail vehicles are all processed, update the corresponding numbers in any consist
        # operation instructions i.e. for coupling or uncoupling
        for cons_rvs in citem.findall(
                'Driver/cDriver/DriverInstructionContainer/cDriverInstructionContainer/DriverInstruction/'
                'cConsistOperations/DeltaTarget/cDriverInstructionTarget/RailVehicleNumber'):
            # Iterate through driver consist instructions and update changed vehicle numbers in the consist
            for crv in cons_rvs.findall('e'):
                for rvp in rv_pairs:
                    if crv.text == rvp[0]:
                        crv.text = rvp[1]
    # All necessary elements processed, now close progress bar window and return the new xml tree object
    progress_win.close()
    return parser_tree


def vehicle_replacer(provider, product, blueprint, name, number, loaded):
    # Check the rail vehicle found by the XML parser against each of the enabled substitutions.
    # A soon as a replacement is made, return to the XML parser and search for the next vehicle.
    if values['Replace_Mk1']:
        mk1_replace(provider, product, blueprint, name, number)
    if values['Replace_Mk2ac'] and mk2ac_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_Mk2df'] and mk2df_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_FSA'] and fsafta_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_HAA'] and haa_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_HTO'] and coal21_t_hto_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_HTV'] and coal21_t_htv_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_VDA'] and vda_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_IHH'] and ihh_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_User'] and user_replace(provider, product, blueprint, name):
        return True
    if values['Replace_HST'] and hst_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C91'] and c91_replace(provider, product, blueprint, name):
        return True
    if values['Replace_C101'] and c101_replace(provider, product, blueprint, name):
        return True
    if values['Replace_C156'] and c156_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C158'] and c158_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C31'] and c31_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C37'] and c37_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C40'] and c40_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C47'] and c47_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C50'] and c50_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C56'] and c56_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C66'] and c66_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C67'] and c67_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C68'] and c68_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C86'] and c86_replace(provider, product, blueprint, name, number):
        return True
    return True


def fix_short_tags(xml_string):
    # This clumsy fix is necessary because sometimes TS requires short xml empty tags and sometimes long ones.
    # The following substitutions should take care of the important exceptions to the long tag default.
    xml_string = re.sub(r'(<cEngineSimContainer.d:id="[\-0-9]*")></cEngineSimContainer>', r'\1/>', xml_string,
                        flags=re.IGNORECASE)
    xml_string = re.sub(r'(<RailVehicleNumber)></RailVehicleNumber>', r'\1/>', xml_string, flags=re.IGNORECASE)
    xml_string = re.sub(r'(<Other)></Other>', r'\1/>', xml_string, flags=re.IGNORECASE)
    xml_string = re.sub(r'(<DeltaTarget)></DeltaTarget>', r'\1/>', xml_string, flags=re.IGNORECASE)
    xml_string = re.sub(r'(<d:nil)></d:nil>', r'\1/>', xml_string, flags=re.IGNORECASE)
    xml_string = re.sub(r'(<DriverInstruction)></DriverInstruction>', r'\1/>', xml_string, flags=re.IGNORECASE)
    xml_string = re.sub(r'(<InitialLevel)></InitialLevel>', r'\1/>', xml_string, flags=re.IGNORECASE)
    xml_string = re.sub(r'(<StaticChildrenMatrix)></StaticChildrenMatrix>', r'\1/>', xml_string, flags=re.IGNORECASE)
    xml_string = re.sub(r'(<RailVehicles)></RailVehicles>', r'\1/>', xml_string, flags=re.IGNORECASE)
    return xml_string


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
    htmrv = "<h1>Rail vehicle list</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n" \
            "    <tr style=\"text-align: right;\">\n      <th>Consist</th>\n      <th>Provider</th>\n" \
            "      <th>Product</th>\n      <th>Blueprint</th>\n      <th>Name</th>\n      <th>Number</th>\n" \
            "      <th>Loaded</th>\n    </tr>\n  </thead>\n  <tbody>\n"
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
        if Path(railworks_path, 'Assets', row[1], row[2], row[3]).is_file():
            tdstyle = ''
        else:
            tdstyle = ' class="missing"'
        for col in row[0:7]:
            col_no += 1
            if rowspan > 0 and col_no == 1:
                col_htm = col_htm + '      <td rowspan=' + str(rowspan) + '><i>' + row[7] + '</i></td>\n'
            elif col_no > 1:
                col_htm = col_htm + '      <td' + tdstyle + '>' + col + '</td>\n'
        col_no = 0
        if (int(row[0]) % 2) == 0:
            htmrv = htmrv + '    <tr>\n' + col_htm + '    </tr>\n'
        else:
            htmrv = htmrv + '    <tr class=\"shaded_row\">\n' + col_htm + '    </tr>\n'
        last_cons = int(row[0])
    htmrv = htmrv + '  </tbody>\n</table>\n<h3>' + str(len(vehicle_list)) + ' vehicles in total in this scenario.</h3>'
    htmas = '\n<h1>List of rail vehicle assets used</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n' \
            '    <tr style=\"text-align: right;\">\n      <th>Provider</th>\n      <th>Product</th>\n    </tr>\n' \
            '  </thead>\n  <tbody>\n'
    unique_assets.sort(key=lambda x: (x[0], x[1]))
    for asset in unique_assets:
        htmas = htmas + '    <tr>\n      <td>' + asset[0] + '</td>\n      <td>' + asset[1] + '</td>\n    </tr>\n'
    htmas = htmas + '  </tbody>\n</table>\n'
    htm = htmhead + htmas + htmrv + '</body>\n</html>\n'
    html_file_path.touch()
    html_file_path.write_text(htm)
    return True


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
                     'Version 0.8a',
                     'Copyright 2021 JR McKenzie', 'https://github.com/jrmckenzie/RSSwapTool')
        elif event == 'Settings':
            if not config.has_section('defaults'):
                config.add_section('defaults')
            config.set('defaults', 'replace_mk1', str(values['Replace_Mk1']))
            config.set('defaults', 'replace_mk2ac', str(values['Replace_Mk2ac']))
            config.set('defaults', 'replace_mk2df', str(values['Replace_Mk2df']))
            config.set('defaults', 'replace_fsa', str(values['Replace_FSA']))
            config.set('defaults', 'replace_haa', str(values['Replace_HAA']))
            config.set('defaults', 'replace_hto', str(values['Replace_HTO']))
            config.set('defaults', 'replace_htv', str(values['Replace_HTV']))
            config.set('defaults', 'replace_vda', str(values['Replace_VDA']))
            config.set('defaults', 'replace_ihh', str(values['Replace_IHH']))
            config.set('defaults', 'replace_user', str(values['Replace_User']))
            config.set('defaults', 'replace_c31', str(values['Replace_C31']))
            config.set('defaults', 'replace_c37', str(values['Replace_C37']))
            config.set('defaults', 'replace_c40', str(values['Replace_C40']))
            config.set('defaults', 'replace_c47', str(values['Replace_C47']))
            config.set('defaults', 'replace_c50', str(values['Replace_C50']))
            config.set('defaults', 'replace_c56', str(values['Replace_C56']))
            config.set('defaults', 'replace_c66', str(values['Replace_C66']))
            config.set('defaults', 'replace_c67', str(values['Replace_C67']))
            config.set('defaults', 'replace_c68', str(values['Replace_C68']))
            config.set('defaults', 'replace_c86', str(values['Replace_C86']))
            config.set('defaults', 'replace_hst', str(values['Replace_HST']))
            config.set('defaults', 'replace_c91', str(values['Replace_C91']))
            config.set('defaults', 'replace_c101', str(values['Replace_C101']))
            config.set('defaults', 'replace_c156', str(values['Replace_C156']))
            config.set('defaults', 'replace_c158', str(values['Replace_C158']))
            with open(path_to_config, 'w') as configfile:
                config.write(configfile)
                configfile.close()
            c86_hc = config.get('defaults', 'c86_hc')
            c56_rf = config.get('defaults', 'c56_rf')
            # The settings button has been pressed, so allow the user to change the RailWorks folder setting
            loclayout = [
                [sg.Text('Settings', justification='c')],
                [sg.Text('Path to RailWorks folder:'),
                 sg.Input(default_text=str(railworks_path), key='RWloc', readonly=True),
                 sg.FolderBrowse(key='RWloc')],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text(
                    'If a BR Blue Class 86 with headcode blinds from the RSC Class 86 Pack is found, you can have it '
                    'replaced with the Vulcan Productions headcode blinds addition to the AP Enhancement Pack if you '
                    'like, or you can replace it with the AP loco with plated over headcode box and marker lights, '
                    'or you can leave the original RSC Class 86 pack in place.',
                    size=(72, 0))],
                [sg.Combo(c86_opts, auto_size_text=True, default_value=c86_hc, key='c86_hc', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text(
                    'If a Railfreight Sectors Class 56 Pack is found, and the depot plaque and/or sector logo of the '
                    'original is not available in the AP Enhancement Pack, you can either replace it with the nearest '
                    'numbered loco in the AP numbering database (depot / sectors will be different) or skip swapping '
                    'and keep the original RSC Railfreight Sectors Class 56.',
                    size=(72, 0))],
                [sg.Combo(c56_opts, auto_size_text=True, default_value=c56_rf, key='c56_rf', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Checkbox('Save a list of all vehicles in the scenario (useful for debugging)', key='save_report',
                             default=config.getboolean('defaults', 'save_report'), enable_events=True,
                             tooltip='This option will save a report listing all the rail vehicles (and their numbers)'
                                     ' in the scenario, in .html format, alongside the scenario output file.')],
                [sg.Button('Save changes'), sg.Button('Cancel')]
            ]
            locwindow = sg.Window('RSSwapTool - Settings', loclayout)
            while True:
                levent, lvalues = locwindow.read()
                if levent == 'Cancel' or sg.WIN_CLOSED:
                    break
                elif levent == 'Save changes':
                    railworks_path = lvalues['RWloc']
                    if not config.has_section('RailWorks'):
                        config.add_section('RailWorks')
                    config.set('RailWorks', 'path', lvalues['RWloc'])
                    if not config.has_section('defaults'):
                        config.add_section('defaults')
                    config.set('defaults', 'c86_hc', str(lvalues['c86_hc']))
                    config.set('defaults', 'c56_rf', str(lvalues['c56_rf']))
                    config.set('defaults', 'save_report', str(lvalues['save_report']))
                    with open(path_to_config, 'w') as configfile:
                        config.write(configfile)
                        configfile.close()
                    railworks_path = Path(railworks_path)
                    break
            locwindow.close()
        elif event == 'Replace!':
            # The replace button has been pressed
            if not config.has_section('defaults'):
                config.add_section('defaults')
            config.set('defaults', 'replace_mk1', str(values['Replace_Mk1']))
            config.set('defaults', 'replace_mk2ac', str(values['Replace_Mk2ac']))
            config.set('defaults', 'replace_mk2df', str(values['Replace_Mk2df']))
            config.set('defaults', 'replace_fsa', str(values['Replace_FSA']))
            config.set('defaults', 'replace_haa', str(values['Replace_HAA']))
            config.set('defaults', 'replace_hto', str(values['Replace_HTO']))
            config.set('defaults', 'replace_htv', str(values['Replace_HTV']))
            config.set('defaults', 'replace_vda', str(values['Replace_VDA']))
            config.set('defaults', 'replace_ihh', str(values['Replace_IHH']))
            config.set('defaults', 'replace_user', str(values['Replace_User']))
            config.set('defaults', 'replace_c31', str(values['Replace_C31']))
            config.set('defaults', 'replace_c37', str(values['Replace_C37']))
            config.set('defaults', 'replace_c40', str(values['Replace_C40']))
            config.set('defaults', 'replace_c47', str(values['Replace_C47']))
            config.set('defaults', 'replace_c50', str(values['Replace_C50']))
            config.set('defaults', 'replace_c56', str(values['Replace_C56']))
            config.set('defaults', 'replace_c66', str(values['Replace_C66']))
            config.set('defaults', 'replace_c67', str(values['Replace_C67']))
            config.set('defaults', 'replace_c68', str(values['Replace_C68']))
            config.set('defaults', 'replace_c86', str(values['Replace_C86']))
            config.set('defaults', 'replace_hst', str(values['Replace_HST']))
            config.set('defaults', 'replace_c91', str(values['Replace_C91']))
            config.set('defaults', 'replace_c101', str(values['Replace_C101']))
            config.set('defaults', 'replace_c156', str(values['Replace_C156']))
            config.set('defaults', 'replace_c158', str(values['Replace_C158']))
            with open(path_to_config, 'w') as configfile:
                config.write(configfile)
                configfile.close()
            if len(values['Scenario_xml']) < 1:
                sg.popup('No scenario selected!')
            else:
                scenarioPath = Path(values['Scenario_xml'])
                outPathStem = scenarioPath.parent / Path(str(scenarioPath.stem) + '-' + time.strftime('%Y%m%d-%H%M%S'))
                inFile = scenarioPath
                cmd = railworks_path / Path('serz.exe')
                serz_output = ''
                vehicle_list = []
                if str(scenarioPath.suffix) == '.bin':
                    # This is a bin file so we need to run serz.exe command to convert it to a readable .xml
                    # intermediate file
                    if not cmd.is_file():
                        sg.popup('serz.exe could not be found in ' + str(railworks_path) + '. Is this definitely your '
                                                                                           'RailWorks folder?',
                                 'This application will now exit.')
                        sys.exit()
                    inFile = scenarioPath.parent / Path(str(scenarioPath.stem) + '.xml')
                    p1 = subprocess.Popen([str(cmd), str(scenarioPath), '/xml:' + str(inFile)], stdout=subprocess.PIPE)
                    p1.wait()
                    serz_output = 'serz.exe ' + p1.communicate()[0].decode('ascii')
                    # Now the intermediate .xml has been created by serz.exe, read it in to this script and do the
                    # processing
                tree = parse_xml(inFile)
                if tree is False:
                    continue
                # Back up the original scenario file, fix some peculiarities of the train simulator .xml,
                # write the final xml out to another temporary file so that serz.exe can convert it back to
                # a .bin file in place of the original.
                scenarioPath.rename(Path(str(outPathStem) + str(scenarioPath.suffix)))
                xmlString = ET.tostring(tree.getroot(), encoding='utf-8', xml_declaration=True,
                                        short_empty_elements=False, method='xml')
                xmlString = fix_short_tags(xmlString.decode())
                xmlFile = scenarioPath.parent / Path(str(scenarioPath.stem) + '.xml')
                xmlFile.touch()
                xmlFile.write_text(xmlString)
                output_message = 'Scenario converted and saved to ' + str(xmlFile)
                html_report_status_text = ''
                if str(scenarioPath.suffix) == '.bin':
                    # Run the serz.exe command again to generate the output scenario .bin file
                    binFile = scenarioPath.parent / Path(str(scenarioPath.stem) + '.bin')
                    p2 = subprocess.Popen([str(cmd), str(xmlFile), '/bin:' + str(binFile)], stdout=subprocess.PIPE)
                    p2.wait()
                    inFile.unlink()
                    output_message = serz_output + '\nserz.exe ' + p2.communicate()[0].decode('ascii')
                output_message = output_message + \
                                 '\nOriginal scenario backup located in ' + str(outPathStem) + str(scenarioPath.suffix)
                if config.getboolean('defaults', 'save_report'):
                    html_report_file = scenarioPath.parent / Path(str(scenarioPath.stem) + '-railvehicle_report.html')
                    convert_vlist_to_html_table(html_report_file)
                    html_report_status_text = 'Report listing all rail vehicles located in ' + str(html_report_file)
                    browser = sg.popup_yes_no(output_message, html_report_status_text,
                                              'Do you want to open the report in your web browser now?')
                    if browser == 'Yes':
                        webbrowser.open(html_report_file)
                else:
                    sg.popup(output_message)
