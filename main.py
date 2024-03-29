#!/usr/bin/python3
#
#     RSSwapTool - A script to swap in up to date or enhanced rolling stock
#     for older versions of stock in Train Simulator scenarios.
#     Copyright (C) 2023 James McKenzie jrmknz@yahoo.co.uk
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
import platform
import configparser
import PySimpleGUI as sg
import webbrowser
from pathlib import Path
from pathlib import PureWindowsPath
from data_file import hha_e_wagons, hha_l_wagons, HTO_141_numbers, HTO_143_numbers, \
    HTO_146_numbers, HTO_rebodied_numbers, HTV_146_numbers, HTV_rebodied_numbers, c158_s9bl_rr, c158_s9bl_nr, \
    c158_s9bl_fgw, c158_s9bl_tpe, c158_s9bl_swt, c158_nwc, c158_dtg_fc, c158_livman_rr, ap40headcodes_69_77, \
    ap40headcodes_62_69, rsc20headcodes_62_69, c168_chiltern, c170_ex_ar_aga_ap, c170_lm, c170_ct_xc, c170_ar23, \
    c170_scotrail, c170_ftpe, c170_ga_hull, c170_mml, c171_southern, c350_lb_ftpe, c365_ecmls_nse, c365_apcxse, \
    c450_gu_swt, c465_se, c375_dtg_pack, c377_lb_se, c377_lg_sn, c377_fcc, c170_bp_name_lookup, c375_dmos_lookup, \
    c456_nse, c456_southern, c319_dest, c350_lm_wcmls, c375_southern_wcmls, c86_TOPS_HC, c350_lm_cc

# If you want to run this script on Linux you must enter the path to the wine executable. You need wine in order to
# run the serz.exe utility which converts between .bin and .xml scenario files.
# If you're not running this script on Linux this line should be left as the default.
wine_executable = '/usr/bin/wine'

# Initialise the script, set the look and feel and get the configuration
version_number = '1.0.10'
version_date = '6 August 2023'
rv_list = []
rv_pairs = []
output_vehicle_list = []
input_vehicle_list = []
layout = []
values = {}
vehicle_db = {}
user_db = {}
vp_blue_47_db = {}
mu_last = 'none'
mso_num = ''
railworks_path = ''
c56_opts = ['Use nearest numbered AP enhanced loco', 'Retain original loco if no matching AP plaque / sector available']
c86_opts = ['Use VP headcode blinds', 'Use AP plated box with markers', 'Do not swap this loco']
fsafta_opts = ['RFD / 2000 era', 'FL / 2000 era', 'FL / 2010 era', 'FL / 2020 era']
fsafta_cube = ['No high cube containers', 'Allow high cube containers']
mgr_types = ['HAA only', 'HCA (canopy) only', 'HFA (canopy) only', 'HAA and HCA (canopy) mixed',
             'HAA, HCA (canopy) and HFA (canopy) mixed', 'HDA only', 'HBA (canopy) only', 'HDA and HBA (canopy) mixed',
             'HMA only', 'HNA (canopy) only', 'HMA and HNA (canopy) mixed', 'Completely random']
mgr_liveries = ['Maroon only', 'Blue only', 'Maroon and Blue', 'Sectors only', 'Sectors and Maroon',
                'Sectors and Blue', 'Completely random']
vda_liveries = ['Maroon only', 'Railfreight only', 'Mostly Maroon', 'Mostly Railfreight', 'Evenly Mixed']
vda_whiteroof_probabilities = ['0', '1', '5', '10', '20', '30', '40', '50', '60', '70', '80', '90', '100']
dirty_probabilities = ['0', '10', '20', '30', '40', '50', '60', '70', '80', '90', '100']
htx_eras = ['Pre-TOPS only', 'Mixed', 'TOPS only']
tail_opts = ['Flashing', 'Steady']
report_opts = ['Don\'t save a report', 'Save details of processed scenario only',
               'Save details of original and processed scenarios']
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
script_path = Path(os.path.abspath(os.path.dirname(sys.argv[0])))
path_to_config = script_path / 'config.ini'
config.read(path_to_config)
# Read configuration and find location of RailWorks folder, or ask user to set it
if config.has_option('RailWorks', 'path'):
    railworks_path = config.get('RailWorks', 'path')
else:
    loclayout = [[sg.T('')],
                 [sg.Text('Please locate your RailWorks folder:'), sg.Input(key='-IN2-', change_submits=False,
                                                                            readonly=True),
                  sg.FolderBrowse(key='RWloc')], [sg.Button('Save')]]
    locwindow = sg.Window('Configure path to RailWorks folder', loclayout)
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
        elif event == 'Save':
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
    config.set('defaults', 'replace_mk3ab', 'True')
    config.set('defaults', 'replace_fsa', 'True')
    config.set('defaults', 'replace_haa', 'True')
    config.set('defaults', 'replace_hha', 'True')
    config.set('defaults', 'replace_hto', 'True')
    config.set('defaults', 'replace_htv', 'True')
    config.set('defaults', 'replace_tta', 'True')
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
    config.set('defaults', 'replace_c87', 'True')
    config.set('defaults', 'replace_hst', 'True')
    config.set('defaults', 'replace_c91', 'False')
    config.set('defaults', 'replace_c101', 'False')
    config.set('defaults', 'replace_c150', 'False')
    config.set('defaults', 'replace_c156', 'False')
    config.set('defaults', 'replace_c158', 'False')
    config.set('defaults', 'replace_c170', 'False')
    config.set('defaults', 'replace_c175', 'False')
    config.set('defaults', 'replace_c221', 'False')
    config.set('defaults', 'replace_c319', 'False')
    config.set('defaults', 'replace_c325', 'False')
    config.set('defaults', 'replace_c350', 'False')
    config.set('defaults', 'replace_c365', 'False')
    config.set('defaults', 'replace_c375', 'False')
    config.set('defaults', 'replace_c450', 'False')
    config.set('defaults', 'replace_c456', 'False')
    config.set('defaults', 'replace_c465', 'False')
    config.set('defaults', 'save_report', report_opts[0])
    config.set('defaults', 'c56_rf', c56_opts[0])
    config.set('defaults', 'c86_hc', c86_opts[0])
    config.set('defaults', 'fsafta_variant', fsafta_opts[0])
    config.set('defaults', 'fsafta_hc', fsafta_cube[0])
    config.set('defaults', 'mgr_type', mgr_types[0])
    config.set('defaults', 'mgr_livery', mgr_liveries[0])
    config.set('defaults', 'vda_livery', vda_liveries[0])
    config.set('defaults', 'vda_whiteroof_probability', vda_whiteroof_probabilities[0])
    config.set('defaults', 'vda_dirty_probability', dirty_probabilities[8])
    config.set('defaults', 'htx_dirty_probability', dirty_probabilities[8])
    config.set('defaults', 'htx_era', htx_eras[1])
    config.set('defaults', 'tail_style', tail_opts[0])
    with open(path_to_config, 'w') as iconfigfile:
        config.write(iconfigfile)
        iconfigfile.close()


def get_my_config_boolean(section, configvalue):
    if config.has_option(section, configvalue):
        return config.getboolean(section, configvalue)
    else:
        config.set(section, configvalue, 'False')
        with open(path_to_config, 'w') as myconfigfile:
            config.write(myconfigfile)
            myconfigfile.close()
        return config.getboolean(section, configvalue)


def get_destination(this_dict, this_key, this_blank):
    if this_key in this_dict:
        return this_dict[this_key]
    else:
        return this_blank


def import_data_from_csv(csv_filename):
    try:
        with open((script_path / csv_filename), 'r') as csv_file:
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
user_db_path = script_path / 'tables/User.csv'
if not user_db_path.is_file():
    head = 'Label,Provider,Product,Blueprint,ReplaceProvider,ReplaceProduct,ReplaceBlueprint,ReplaceName,NumbersDcsv\n'
    user_db_path.touch()
    user_db_path.write_text(head)
vehicle_db = import_data_from_csv('tables/Replacements.csv')
user_db = import_data_from_csv('tables/User.csv')
vp_blue_47_db = import_data_from_csv('tables/Class47BRBlue_numbers.csv')

# Set the layout of the GUI
left_column = [
    [sg.Text('RSSwapTool', font='Helvetica 16'), sg.Text('v' + version_number, font='Helvetica 8')],
    [sg.Text('© 2023 JR McKenzie', font='Helvetica 7')],
    [sg.FileBrowse('Select scenario file to process', key='Scenario_xml', tooltip='Locate the scenario .bin or .xml '
                                                                                  'file you wish to process')],
    [sg.Text('Tick the boxes below to choose the\nsubstitutions you would like to make.')],
    [sg.Checkbox('Replace Mk1 coaches', default=get_my_config_boolean('defaults', 'replace_mk1'), enable_events=True,
                 tooltip='Tick to enable replacing of Mk1 coaches with AP Mk1 Coach Pack Vol. 1',
                 key='Replace_Mk1')],
    [sg.Checkbox('Replace Mk2A-C coaches', default=get_my_config_boolean('defaults', 'replace_mk2ac'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Mk2a coaches with AP Mk2A-C Pack', key='Replace_Mk2ac')],
    [sg.Checkbox('Replace Mk2D-F coaches', default=get_my_config_boolean('defaults', 'replace_mk2df'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Mk2e coaches with AP Mk2D-F Pack', key='Replace_Mk2df')],
    [sg.Checkbox('Replace Mk3A-B coaches', default=get_my_config_boolean('defaults', 'replace_mk3ab'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Mk3A-B coaches with AP Mk3A-B Pack', key='Replace_Mk3ab')],
    [sg.Checkbox('Replace FSA/FTA wagons', default=get_my_config_boolean('defaults', 'replace_fsa'), enable_events=True,
                 tooltip='Tick to enable replacing of FSA wagons with AP FSA/FTA Wagon Pack', key='Replace_FSA')],
    [sg.Checkbox('Replace HAA wagons', default=get_my_config_boolean('defaults', 'replace_haa'), enable_events=True,
                 tooltip='Tick to enable replacing of HAA wagons with AP MGR Wagon Pack', key='Replace_HAA')],
    [sg.Checkbox('Replace HHA wagons', default=get_my_config_boolean('defaults', 'replace_hha'), enable_events=True,
                 tooltip='Tick to enable replacing of HHA wagons with AP HHA Wagon Pack', key='Replace_HHA')],
    [sg.Checkbox('Replace unfitted 21t coal wagons', default=get_my_config_boolean('defaults', 'replace_hto'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of unfitted 21t coal wagons with Fastline Simulation HTO wagons',
                 key='Replace_HTO')],
    [sg.Checkbox('Replace fitted 21t coal wagons', default=get_my_config_boolean('defaults', 'replace_htv'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of fitted 21t coal wagons with Fastline Simulation HTV wagons',
                 key='Replace_HTV')],
    [sg.Checkbox('Replace TTA wagons', default=get_my_config_boolean('defaults', 'replace_tta'), enable_events=True,
                 tooltip='Tick to enable replacing of TTA wagons with Armstrong Powerhouse TTA pack',
                 key='Replace_TTA')],
    [sg.Checkbox('Replace VDA wagons', default=get_my_config_boolean('defaults', 'replace_vda'), enable_events=True,
                 tooltip='Tick to enable replacing of VDA wagons with Fastline Simulation VDA pack',
                 key='Replace_VDA')],
    [sg.Checkbox('Replace IHH stock', default=get_my_config_boolean('defaults', 'replace_ihh'), enable_events=True,
                 tooltip='Tick to enable replacing of old Iron Horse House (IHH) stock, if your scenario contains any'
                         ' (if in doubt, leave this unticked)',
                 key='Replace_IHH')],
]
mid_column = [
    [sg.Checkbox('Replace User-configured stock', default=get_my_config_boolean('defaults', 'replace_user'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of user-configured stock, contained in file User.csv '
                         '(leave this unticked unless you have added your own substitutions to User.csv).',
                 key='Replace_User')],
    [sg.Checkbox('Replace Class 31s', default=get_my_config_boolean('defaults', 'replace_c31'), enable_events=True,
                 tooltip='Replace Class 31s with AP enhancement pack equivalent', key='Replace_C31')],
    [sg.Checkbox('Replace Class 37s', default=get_my_config_boolean('defaults', 'replace_c37'), enable_events=True,
                 tooltip='Replace Class 37s with AP equivalent', key='Replace_C37')],
    [sg.Checkbox('Replace Class 40s', default=get_my_config_boolean('defaults', 'replace_c40'), enable_events=True,
                 tooltip='Replace DT Class 40s with AP/RailRight equivalent', key='Replace_C40')],
    [sg.Checkbox('Replace Class 47s', default=get_my_config_boolean('defaults', 'replace_c47'), enable_events=True,
                 tooltip='Replace BR Blue Class 47s with Vulcan Productions BR Blue Class 47 Pack versions',
                 key='Replace_C47')],
    [sg.Checkbox('Replace Class 50s', default=get_my_config_boolean('defaults', 'replace_c50'), enable_events=True,
                 tooltip='Replace MeshTools Class 50s with AP equivalent', key='Replace_C50')],
    [sg.Checkbox('Replace Class 56s', default=get_my_config_boolean('defaults', 'replace_c56'), enable_events=True,
                 tooltip='Replace RSC Class 56 Railfreight Sectors with AP enhancement pack equivalent',
                 key='Replace_C56')],
    [sg.Checkbox('Replace Class 66s', default=get_my_config_boolean('defaults', 'replace_c66'), enable_events=True,
                 tooltip='Replace Class 66s with AP enhancement pack equivalent', key='Replace_C66')],
    [sg.Checkbox('Replace Class 67s', default=get_my_config_boolean('defaults', 'replace_c67'), enable_events=True,
                 tooltip='Replace Class 67s with AP enhancement pack equivalent', key='Replace_C67')],
    [sg.Checkbox('Replace Class 68s', default=get_my_config_boolean('defaults', 'replace_c68'), enable_events=True,
                 tooltip='Replace Class 68s with AP enhancement pack equivalent', key='Replace_C68')],
    [sg.Checkbox('Replace Class 86s', default=get_my_config_boolean('defaults', 'replace_c86'), enable_events=True,
                 tooltip='Replace Class 86s with AP enhancement pack equivalent', key='Replace_C86')],
    [sg.Checkbox('Replace Class 87s', default=get_my_config_boolean('defaults', 'replace_c87'), enable_events=True,
                 tooltip='Replace Class 87s with AP equivalent', key='Replace_C87')],
    [sg.Checkbox('Replace HST sets', default=get_my_config_boolean('defaults', 'replace_hst'), enable_events=True,
                 tooltip='Tick to enable replacing of HST sets with AP enhanced versions (Valenta, MTU, VP185)',
                 key='Replace_HST')],
    [sg.Checkbox('Replace Class 91 EC sets', default=get_my_config_boolean('defaults', 'replace_c91'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 91 East Coast sets with AP enhanced versions',
                 key='Replace_C91')],
    [sg.Checkbox('Replace Class 101 sets', default=get_my_config_boolean('defaults', 'replace_c101'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of retired RSC Class101Pack with RSC BritishRailClass101 sets',
                 key='Replace_C101')],
    [sg.Checkbox('Replace Class 150/2 sets', default=get_my_config_boolean('defaults', 'replace_c150'),
                 enable_events=True,
                 tooltip='Tick to enable replacing Thomson-Oovee Class 150s with AP Class 150/2', key='Replace_C150')],
]
right_column = [
    [sg.Checkbox('Replace Class 156 sets', default=get_my_config_boolean('defaults', 'replace_c156'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Oovee Class 156s with AP Class 156', key='Replace_C156')],
    [sg.Checkbox('Replace Class 158, 159 sets', default=get_my_config_boolean('defaults', 'replace_c158'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of North Wales Coast / Settle Carlisle / Fife Circle Class 158s '
                         'with AP enhanced versions (Cummins, Perkins)',
                 key='Replace_C158')],
    [sg.Checkbox('Replace Class 168, 170, 171 sets', default=get_my_config_boolean('defaults', 'replace_c170'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 168s / 170s with AP enhanced versions',
                 key='Replace_C170')],
    [sg.Checkbox('Replace Class 175 sets', default=get_my_config_boolean('defaults', 'replace_c175'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 175s with AP enhanced versions',
                 key='Replace_C175')],
    [sg.Checkbox('Replace Class 220, 221 sets', default=get_my_config_boolean('defaults', 'replace_c221'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of DTG Voyager with JT Advanced Voyager',
                 key='Replace_C221')],
    [sg.Checkbox('Replace Class 319 sets', default=get_my_config_boolean('defaults', 'replace_c319'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 319s with AP versions',
                 key='Replace_C319')],
    [sg.Checkbox('Replace Class 325 sets', default=get_my_config_boolean('defaults', 'replace_c325'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 325s with AP enhanced versions',
                 key='Replace_C325')],
    [sg.Checkbox('Replace Class 350 sets', default=get_my_config_boolean('defaults', 'replace_c350'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 350s with AP enhanced versions',
                 key='Replace_C350')],
    [sg.Checkbox('Replace Class 365 sets', default=get_my_config_boolean('defaults', 'replace_c365'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 365s with AP enhanced versions',
                 key='Replace_C365')],
    [sg.Checkbox('Replace Class 375/377/379/387 sets', default=get_my_config_boolean('defaults', 'replace_c375'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 375/377/379/387 with AP enhanced versions',
                 key='Replace_C375')],
    [sg.Checkbox('Replace Class 444, 450 sets', default=get_my_config_boolean('defaults', 'replace_c450'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 444, 450s with AP enhanced versions',
                 key='Replace_C450')],
    [sg.Checkbox('Replace Class 456 sets', default=get_my_config_boolean('defaults', 'replace_c456'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 456s with AP/Waggonz versions',
                 key='Replace_C456')],
    [sg.Checkbox('Replace Class 465 sets', default=get_my_config_boolean('defaults', 'replace_c465'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Class 465s with AP enhanced versions',
                 key='Replace_C465')],
    [sg.Button('Replace!'), sg.Button('Settings'), sg.Button('About'), sg.Button('Exit')],
]

# Set the layout of the application window
layout = [
    [
        sg.Column(left_column),
        sg.VSeperator(),
        sg.Column(mid_column),
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
                # This dcsv number is already in use in another swapped loco - move on and try the next one
                continue
            if ithis_rv > dcsv_nm:
                # This dcsv number is still less than the number we're looking for - but remember how close it is and try the
                # next one to see if it is a match or is even further away than this number.
                diff = ithis_rv - dcsv_nm
            elif ithis_rv == dcsv_nm:
                # A matching number is available and has been found - use it.
                return str(curr_nm)
            elif ithis_rv < dcsv_nm:
                # We have overshot the number we are looking for - but if this number is even further from the number
                # we're looking for than the last one we looked at, use the last one.
                if dcsv_nm - ithis_rv > diff:
                    return str(last_nm)
                else:
                    # We've checked and even though we have overshot the number we are looking for, this number is
                    # closer to the number we're looking for than the last one we looked at - so use it.
                    return str(curr_nm)
            last_nm = curr_nm
    # We didn't find a number to use. We must have reached the end of the available numbers so we will have to use the
    # last available number we found.
    return str(last_nm)


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


def alternate_mu_driving_vehicles(this_bp, this_v_type, vehicle_a, vehicle_b):
    # This function makes sure that driving vehicles in a multiple unit consist alternate between DxxA and DxxB
    # so that the AP scripting of multiple unit sets works properly.
    global mu_last
    if this_v_type.upper() == vehicle_a.upper():
        # Test if last Driving vehicle was a DxxA - if so, swap this one for a DxxB
        if mu_last == vehicle_a:
            # Swap for a DxxB
            this_bp = re.sub(vehicle_a, vehicle_b, this_bp, flags=re.IGNORECASE)
            # Remember that the last driven vehicle in this consist was a DxxB.
            mu_last = vehicle_b
        else:
            # Leave the DxxA as it is. Remember that the last driving vehicle processed in this
            # consist was a DxxA.
            mu_last = vehicle_a
        return this_bp
    elif this_v_type.upper() == vehicle_b.upper():
        # Test if last Driving vehicle was a DxxB - if so, swap this one for a DxxA
        if mu_last == vehicle_b:
            # Swap for a DxxA
            this_bp = re.sub(vehicle_b, vehicle_a, this_bp, flags=re.IGNORECASE)
            # Remember that the last cab vehicle in this consist was a DxxA.
            mu_last = vehicle_a
        else:
            # Leave the DxxB as it is. Remember that the last driving vehicle processed in this
            # consist was a DxxB.
            mu_last = vehicle_b
    return this_bp


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
    rv_digits = int(re.sub('[^0-9]', "", this_rv))
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


def get_ap_name_from_bp(this_vehicle_db, this_bp):
    for i in range(0, len(this_vehicle_db)):
        if this_bp.upper() == this_vehicle_db[i][5].upper():
            return this_vehicle_db[i][6]
    return False


def direction_flip(flipped, followers):
    if flipped.text == '0':
        flipped.text = '1'
    else:
        flipped.text = '0'
    for follower in followers.findall('Network-cTrackFollower'):
        dir = follower.find('Direction/Network-cDirection/_dir')
        if dir.text == 'forwards':
            dir.text = 'backwards'
        else:
            dir.text = 'forwards'
    return True


def add_taillamp(tailmarker, this_blueprint, bp_suffix, this_name, name_suffix, flipped, followers):
    # The vehicle is at one end of the consist and should have a red tail light
    this_blueprint = re.sub('\.xml', bp_suffix, this_blueprint, flags=re.IGNORECASE)
    this_name = this_name + name_suffix
    # If the vehicle is at the top of the consist it will need to be flipped to have the tail
    # light facing the right direction
    if tailmarker == 0:
        if flipped.text == '0':
            direction_flip(flipped, followers)
    if tailmarker == 2:
        if flipped.text == '1':
            direction_flip(flipped, followers)
    return this_blueprint, this_name


def haa_replace(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker):
    # Replace HAA wagons
    for i in range(0, len(vehicle_db['HAA'])):
        this_vehicle = vehicle_db['HAA'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = 'AP'
                    product.text = 'HAAWagonPack01'
                    weathering = random.choice([('', ''), ('2', ' W1'), ('2', ' W1'), ('3', ' W2'), ('3', ' W2')])
                    if 'eTrue' in loaded.text:
                        load = ['_LD', 'Loaded']
                    else:
                        load = ['', 'Empty']
                    variant = config.get('defaults', 'mgr_variant', fallback='HAA')
                    livery = config.get('defaults', 'mgr_livery', fallback='Maroon Only')
                    if variant == 'HAA only':
                        bp = 'HAA'
                    elif variant == 'HCA (canopy) only':
                        bp = 'HCA'
                    elif variant == 'HFA (canopy) only':
                        bp = 'HFA'
                    elif variant == 'HAA and HCA (canopy) mixed':
                        bp = random.choice(['HAA', 'HCA'])
                    elif variant == 'HAA, HCA (canopy) and HFA (canopy) mixed':
                        bp = random.choice(['HAA', 'HCA', 'HFA'])
                    elif variant == 'HDA only':
                        bp = 'HDA'
                    elif variant == 'HBA (canopy) only':
                        bp = 'HBA'
                    elif variant == 'HDA and HBA (canopy) mixed':
                        bp = random.choice(['HDA', 'HBA'])
                    elif variant == 'HMA only':
                        bp = 'HMA'
                    elif variant == 'HNA (canopy) only':
                        bp = 'HNA'
                    elif variant == 'HMA and HNA (canopy) mixed':
                        bp = random.choice(['HMA', 'HNA'])
                    else:
                        # Completely random MGR wagon
                        bp = random.choice(['HAA', 'HBA', 'HCA', 'HDA', 'HFA', 'HMA', 'HNA'])
                    if livery == 'Maroon only':
                        lv = ['EWS', ' Red ']
                    elif livery == 'Blue only':
                        lv = ['Blue', ' Blue ']
                    elif livery == 'Maroon and Blue':
                        lv = random.choice([('EWS', ' Red '), ('Blue', ' Blue ')])
                    elif livery == 'Sectors and Maroon':
                        lv = random.choice([('Sector', ' Sector '), ('EWS', ' Red ')])
                    elif livery == 'Sectors and Blue':
                        lv = random.choice([('Sector', ' Sector '), ('Blue', ' Blue ')])
                    else:
                        # Completely random livery
                        lv = random.choice([('EWS', ' Red '), ('Blue', ' Blue '), ('Sector', ' Sector ')])
                    blueprint.text = 'RailVehicles\\Freight\\HAA\\' + lv[0] + weathering[0] + '\\' + bp + load[
                        0] + '.xml'
                    name.text = 'AP ' + bp + lv[1] + load[1] + weathering[1]
                    if not tailmarker == 1:
                        # change to a tail lamp carrying wagon and try to orient it with the lamp outward facing
                        blueprint.text, name.text = add_taillamp(tailmarker, blueprint.text, '_TL.xml', name.text,
                                                                 ' TL', flipped, followers)
                    # Now extract the vehicle number
                    rv_list.append(number.text)
                    return True
    return False


def hha_replace(provider, product, blueprint, name, number, loaded):
    # Replace HHA wagons
    for i in range(0, len(vehicle_db['HHA'])):
        this_vehicle = vehicle_db['HHA'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = 'AP'
                    product.text = 'HHAWagonPack01'
                    # Replace a loaded wagon
                    if 'eTrue' in loaded.text or bool(re.search('_LOADED', blueprint.text, flags=re.IGNORECASE)):
                        idx = random.randrange(0, len(hha_l_wagons))
                        # Select at random one of the wagons in the list of HHA loaded wagons to swap in
                        blueprint.text = hha_l_wagons[idx][0]
                        name.text = hha_l_wagons[idx][1]
                    # Replace an empty wagon
                    else:
                        idx = random.randrange(0, len(hha_e_wagons))
                        # Select at random one of the wagons in the list of HHA empty wagons to swap in
                        blueprint.text = hha_e_wagons[idx][0]
                        name.text = hha_e_wagons[idx][1]
                    # Now extract the vehicle number
                    rv_list.append(number.text)
                    return True
    return False


def fsafta_replace(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker):
    if bool(fsa_replace(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker)):
        return True
    if bool(fta_replace(provider, product, blueprint, name, number, loaded)):
        return True
    return False


def fsa_replace(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker):
    for i in range(0, len(vehicle_db['FSA'])):
        this_vehicle = vehicle_db['FSA'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    era = config.get('defaults', 'fsafta_variant', fallback='FL / 2000 era')
                    if era == 'RFD / 2000 era':
                        w = ['RfD', '', '']
                    elif era == 'FL / 2000 era':
                        w = ['FL', '_2000', '(2000)']
                    elif era == 'FL / 2010 era':
                        w = ['FL', '_2010', '(2010)']
                    else:
                        # FL / 2020 era
                        w = ['FL', '_2020', '(2020)']
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = re.sub('RFD|FL', w[0].upper(), re.sub('_20[0-2]0', w[1], this_vehicle[5]),
                                            flags=re.IGNORECASE)
                    name.text = re.sub('RFD|FL', w[0], re.sub('\(20[0-2]0\)', w[2], this_vehicle[6]),
                                       flags=re.IGNORECASE)
                    rv_orig = number.text
                    if 'eFalse' in loaded.text:
                        # Wagon is unloaded
                        number.text = dcsv_get_num(
                            Path(railworks_path, 'Assets/AP/FSAWagonPack/RailVehicles/Freight/FL/FSA.dcsv'),
                            number.text,
                            '([0-9]{6})(.*)')
                        # Change the blueprint and name to the unloaded wagon
                        blueprint.text = re.sub('FSA[a-zA-Z0-9_]*.xml', 'FSA.xml', blueprint.text, flags=re.IGNORECASE)
                        name.text = re.sub('AP.FSA.([a-zA-Z]*).*', r'AP FSA \1', name.text, flags=re.IGNORECASE)
                    else:
                        # Check if high cube containers are allowed
                        if config.get('defaults', 'fsafta_hc',
                                      fallback='No high cube containers') == 'Allow high cube containers':
                            dcsv = re.sub('_No_HC', '', this_vehicle[7].replace('\\', '/'))
                        else:
                            dcsv = this_vehicle[7].replace('\\', '/')
                        number.text = dcsv_get_num(
                            Path(railworks_path, 'Assets/AP/FSAWagonPack', dcsv), number.text,
                            '([0-9]{6})(.*)')
                    if not tailmarker == 1:
                        # change to a tail lamp carrying wagon and try to orient it with the lamp outward facing
                        blueprint.text, name.text = add_taillamp(tailmarker, blueprint.text, '.xml', name.text,
                                                                 '', flipped, followers)
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
                    era = config.get('defaults', 'fsafta_variant', fallback='FL / 2000 era')
                    if era == 'RFD / 2000 era':
                        w = ['RfD', '', '']
                    elif era == 'FL / 2000 era':
                        w = ['FL', '_2000', '(2000)']
                    elif era == 'FL / 2010 era':
                        w = ['FL', '_2010', '(2010)']
                    else:
                        # FL / 2020 era
                        w = ['FL', '_2020', '(2020)']
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = re.sub('RFD|FL', w[0].upper(), re.sub('_20[0-2]0', w[1], this_vehicle[5]),
                                            flags=re.IGNORECASE)
                    name.text = re.sub('RFD|FL', w[0], re.sub('\(20[0-2]0\)', w[2], this_vehicle[6]),
                                       flags=re.IGNORECASE)
                    rv_orig = number.text
                    if 'eFalse' in loaded.text:
                        # Wagon is unloaded
                        number.text = dcsv_get_num(
                            Path(railworks_path, 'Assets/AP/FSAWagonPack/RailVehicles/Freight/FL/FTA.dcsv'),
                            number.text,
                            '([0-9]{6})(.*)')
                        # Change the blueprint and name to the unloaded wagon
                        blueprint.text = re.sub('FTA[a-zA-Z0-9_]*.xml', 'FTA.xml', blueprint.text, flags=re.IGNORECASE)
                        name.text = re.sub('AP.FTA.([a-zA-Z]*).*', r'AP FTA \1', name.text, flags=re.IGNORECASE)
                    else:
                        # Check if high cube containers are allowed
                        if config.get('defaults', 'fsafta_hc',
                                      fallback='No high cube containers') == 'Allow high cube containers':
                            dcsv = re.sub('_No_HC', '', this_vehicle[7].replace('\\', '/'))
                        else:
                            dcsv = this_vehicle[7].replace('\\', '/')
                        number.text = dcsv_get_num(
                            Path(railworks_path, 'Assets/AP/FSAWagonPack', dcsv), number.text,
                            '([0-9]{6})(.*)')
                    rv_pairs.append([rv_orig, number.text])
                    rv_list.append(number.text)
                    return True
    return False


def tta_replace(provider, product, blueprint, name, number, loaded):
    for i in range(0, len(vehicle_db['TTA'])):
        this_vehicle = vehicle_db['TTA'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    (w_blueprint, w_name) = set_weathering(3, this_vehicle)
                    blueprint.text = w_blueprint
                    name.text = w_name
                    number_suffix = this_vehicle[7]
                    rv_orig = number.text
                    if number_suffix[0:1] == ';' and 'eTrue' in loaded.text:
                        # Wagon is loaded
                        # Change the blueprint and name to the unloaded wagon
                        blueprint.text = re.sub('.xml', '_LD.xml', blueprint.text, flags=re.IGNORECASE)
                        name.text = re.sub('Empty', 'Loaded', name.text, flags=re.IGNORECASE)
                    elif number_suffix[0:1] != ';':
                        number_suffix = number_suffix[1:]
                    loaded.text = 'eFalse'
                    number.text = rv_orig + number_suffix
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
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), num, '([0-9]{4,5})(.*)')
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


def mk3ab_replace(provider, product, blueprint, name, number):
    # Replace any Mk3a/bs - loop through the VehicleDB['Mk3ab'] array of coaches to search for
    for i in range(0, len(vehicle_db['Mk3ab'])):
        this_vehicle = vehicle_db['Mk3ab'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    ap_suffix = this_vehicle[7]
                    rv_orig = number.text
                    # Now extract the region code (if there is one) and the coach number
                    nm = re.search('([a-zA-Z]{0,1})([0-9]{5})', number.text)
                    if nm:
                        reg = nm.group(1).upper()
                        num = nm.group(2)
                        if len(reg) > 0 and 'R=Z' not in ap_suffix:
                            ap_suffix = ap_suffix + ";R=" + reg
                        rv_num = num + ap_suffix
                        rv_pairs.append([rv_orig, rv_num])
                        rv_list.append(rv_num)
                        # Following line sets AP coach number
                        number.text = rv_num
                    return True
    return False


def vda_replace(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker):
    # Replace VDA wagons
    if 'JL' in provider.text:
        if 'WHL' in product.text:
            bp = re.search(re.escape(r'RailVehicles\Freight\VDA\VDA.xml'), blueprint.text, flags=re.IGNORECASE)
            if bp:
                provider.text = 'FastlineSimulation'
                rv_orig = number.text
                if 'eTrue' in loaded.text:
                    load = 'L'  # Replace a loaded wagon
                else:
                    load = 'E'  # Replace an empty wagon
                livery = config.get('defaults', 'vda_livery', fallback='Maroon only')
                if livery == 'Maroon only':
                    lv = 'M'
                elif livery == 'Railfreight only':
                    lv = 'RF'
                elif livery == 'Mostly Maroon':
                    lv = random.choice(['M', 'M', 'M', 'RF'])  # 25% probability of Railfreight
                elif livery == 'Mostly Railfreight':
                    lv = random.choice(['RF', 'RF', 'RF', 'M'])  # 75% probability of Railfreight
                else:
                    # Evenly mixed Maroon and Railfreight vans
                    lv = random.choice(['M', 'RF'])  # 50% probability of Railfreight
                white_probability = int(config.get('defaults', 'vda_whiteroof_probability', fallback='0'))
                white_dicethrow = random.randrange(1, 101)
                dirty_probability = int(config.get('defaults', 'vda_dirty_probability', fallback='90'))
                dirty_dicethrow = random.randrange(1, 101)
                if white_dicethrow <= white_probability:
                    white = 'W'
                else:
                    white = ''
                if dirty_dicethrow <= dirty_probability:
                    weathering = 'D'
                else:
                    weathering = 'C'
                # Choose a random VDA lot and remap the vehicle number to somewhere within the numbers this lot was
                # allocated. 750 vehicles were constructed, 330 in lot 3855, 20 in lot 3890, 100 in lot 3856 and 300
                # in lot 3908.
                x = random.randrange(0, 750)
                rv_num = int(number.text)
                if 0 <= x < 330:
                    lot = '3855'
                    m = rv_num % 330
                elif 330 <= x < 350:
                    lot = '3890'
                    m = 330 + rv_num % 20
                elif 350 <= x < 450:
                    lot = '3856'
                    m = 350 + rv_num % 100
                else:
                    lot = '3908'
                    m = 9450 + rv_num % 300
                rv_num = str(m + 200650) + '#####'
                product.text = 'VDA Vans lot ' + lot
                blueprint.text = 'RailVehicles\\Freight\\VDA Vans\\VDA_' + lot + '_' + lv + white + '_' + load + weathering + '.xml'
                name.text = 'VDA: ' + lv + white + ' ' + lot + ' - ' + load + '.' + weathering
                if not tailmarker == 1:
                    # change to a tail lamp carrying wagon and try to orient it with the lamp outward facing
                    tail_style = config.get('defaults', 'tail_style', fallback='Flashing')
                    tail_bp = 'R.xml'
                    tail_name = 'R'
                    if tail_style == 'Flashing':
                        # Only RailFreight or Dirty Maroon vans can have flashing tail lamps
                        if lv == 'RF' or (lv == 'M' and weathering == 'D'):
                            tail_bp = 'Rb.xml'
                            tail_name = 'Rb'
                    blueprint.text, name.text = add_taillamp(tailmarker, blueprint.text, tail_bp, name.text,
                                                             tail_name, flipped, followers)
                    # If the vehicle is it the top of the consist it will need to be flipped to have the tail
                    # light facing the right direction
                    if tailmarker == 0:
                        if flipped.text == '0':
                            direction_flip(flipped, followers)
                    if tailmarker == 2:
                        if flipped.text == '1':
                            direction_flip(flipped, followers)
                # Now process the vehicle number
                rv_list.append(str(rv_num))
                rv_pairs.append([rv_orig, str(rv_num)])
                # Set Fastline wagon number
                number.text = str(rv_num)
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
                    rv_num = int(number.text.replace('B', ''))
                    if 'eTrue' in loaded.text:
                        load = 'L'  # Replace a loaded wagon
                    else:
                        load = 'E'  # Replace an empty wagon
                    # Choose a random HTO lot and remap the vehicle number to somewhere within the numbers this lot was
                    # allocated. 11708 vehicle numbers are possible: 1200 in lot 141, 2750 in lot 143, 6050 in lot 146,
                    # and 1708 rebodied.
                    x = random.randrange(0, 11708)
                    pretops_only = random.choice(['a', 'b', 'c'])
                    no_pretops = 'd'
                    if 0 <= x < 1200:
                        lot = ['Dia 141', '01']
                        m = HTO_141_numbers[rv_num % 1200]
                        dirty_w = random.choice([('W2_', 'W2.'), ('W_', 'W1.')])
                        clean_w = random.choice([('C2_', 'C2.'), ('C_', 'C1.')])
                    elif 1200 <= x < 3950:
                        lot = ['Dia 143', '02']
                        m = HTO_143_numbers[rv_num % 2750]
                        dirty_w = random.choice([('02_W_', 'W2.'), ('W_', 'W1.')])
                        clean_w = random.choice([('02_C_', 'C2.'), ('C_', 'C1.')])
                    elif 3950 <= x < 10000:
                        lot = ['Dia 146', '04']
                        m = HTO_146_numbers[rv_num % 6050]
                        dirty_w = random.choice([('W2_', 'W2.'), ('W_', 'W.')])
                        clean_w = random.choice([('C2_', 'C2.'), ('C_', 'C.')])
                    else:
                        lot = ['Rebodied', '18']
                        m = HTO_rebodied_numbers[rv_num % 1708]
                        dirty_w = random.choice([('B_W_', 'B.W.'), ('G_W_', 'G.W.')])
                        clean_w = random.choice([('B_C_', 'B.C.'), ('G_C_', 'G.C.')])
                        pretops_only = 'c'
                        no_pretops = random.choice(['a', 'b', 'd'])
                    dirty_probability = int(config.get('defaults', 'htx_dirty_probability', fallback='90'))
                    dirty_dicethrow = random.randrange(1, 101)
                    if dirty_dicethrow <= dirty_probability:
                        weathering = dirty_w
                    else:
                        weathering = clean_w
                    data_paneltypes = config.get('defaults', 'htx_era', fallback='Mixed')
                    if data_paneltypes == 'Pre-TOPS only':
                        rv_prefix = pretops_only
                    elif data_paneltypes == 'TOPS only':
                        rv_prefix = no_pretops
                    else:
                        rv_prefix = ''
                    this_blueprint = 'RailVehicles\\Freight\\HTO\\FS_HT0' + lot[1] + 'A_' + weathering[
                        0] + load + '.xml'
                    this_name = 'HTO 21t Hopper - ' + lot[0] + ': ' + weathering[1] + load
                    rv_num = rv_prefix + 'B' + str(m)
                    product.text = 'HTO 21t Hoppers - ' + lot[0]
                    blueprint.text = this_blueprint
                    name.text = this_name
                    number.text = rv_num
                    rv_list.append(rv_num)
                    rv_pairs.append([rv_orig, rv_num])
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
                    rv_num = int(number.text.replace('B', ''))
                    if 'eTrue' in loaded.text:
                        load = 'L'  # Replace a loaded wagon
                    else:
                        load = 'E'  # Replace an empty wagon
                    # Choose a random HTV lot and remap the vehicle number to somewhere within the numbers this lot was
                    # allocated. 2579 vehicle numbers are possible: 441 in lot 146, 2138 rebodied.
                    x = random.randrange(0, 2579)
                    if 0 <= x < 441:
                        lot = 'Dia 146'
                        m = HTV_146_numbers[rv_num % 441]
                        dirty_w = random.choice([('FS_HT004D_01_W_', 'B1.W.'), ('FS_HT004D_02_W_', 'B2.W.')])
                        clean_w = random.choice([('FS_HT004D_01_C_', 'B1.C.'), ('FS_HT004D_02_C_', 'B2.C.')])
                        pretops_only = random.choice(['a', 'b', 'c'])
                        no_pretops = 'd'
                    else:
                        lot = 'Rebodied'
                        m = HTV_rebodied_numbers[rv_num % 2138]
                        dirty_w = random.choice([('Dirty\\HTV_B1_W_', 'B1.W.'), ('Dirty 2\\HTV_B2_W_', 'B2.W.'),
                                                 ('Dirty 3\\HTV_M_W_', 'M.W.')])
                        clean_w = random.choice([('Clean\\HTV_B1_C_', 'B1.C.'), ('Clean 2\\HTV_B2_C_', 'B2.C.'),
                                                 ('Clean 3\\HTV_M_C_', 'M.C.')])
                        pretops_only = '4'
                        no_pretops = random.choice(['1', '2', '3'])
                    dirty_probability = int(config.get('defaults', 'htx_dirty_probability', fallback='90'))
                    dirty_dicethrow = random.randrange(1, 101)
                    if dirty_dicethrow <= dirty_probability:
                        weathering = dirty_w
                    else:
                        weathering = clean_w
                    data_paneltypes = config.get('defaults', 'htx_era', fallback='Mixed')
                    if data_paneltypes == 'Pre-TOPS only':
                        rv_prefix = pretops_only
                    elif data_paneltypes == 'TOPS only':
                        rv_prefix = no_pretops
                    else:
                        rv_prefix = ''
                    this_blueprint = 'RailVehicles\\Freight\\HTV\\' + weathering[0] + load + '.xml'
                    this_name = 'HTV 21t ' + lot + ': ' + weathering[1] + load
                    rv_num = rv_prefix + 'B' + str(m)
                    product.text = 'HTV 21t Hoppers - ' + lot
                    blueprint.text = this_blueprint
                    name.text = this_name
                    number.text = rv_num
                    rv_list.append(rv_num)
                    rv_pairs.append([rv_orig, rv_num])
                    return True
    return False


def ihh_replace(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker):
    if bool(ihh_bonus_replace(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker)):
        return True
    if bool(ihh_c17_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c20_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c25_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c27_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c40_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c45_replace(provider, product, blueprint, name, number)):
        return True
    if bool(ihh_c56_replace(provider, product, blueprint, name, number)):
        return True
    return False


def ihh_bonus_replace(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker):
    for i in range(0, len(vehicle_db['IHH_Bonus'])):
        this_vehicle = vehicle_db['IHH_Bonus'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text
                    guv = re.search('guv', blueprint.text, flags=re.IGNORECASE)
                    cao = re.search('20t', blueprint.text, flags=re.IGNORECASE)
                    hea = re.search('hea railfreight', blueprint.text, flags=re.IGNORECASE)
                    mcv = re.search('16tmineralwagon', blueprint.text, flags=re.IGNORECASE)
                    tip = re.search('iron ore tippler', blueprint.text, flags=re.IGNORECASE)
                    c47 = re.search('brush_4_bue', blueprint.text, flags=re.IGNORECASE)
                    if guv:
                        if not bool(re.match('[A-Z][0-9]{5}', number.text)):
                            # If the number is not in expected format, choose a random one.
                            number.text = 'M' + str(random.randint(86078, 86984))
                    elif cao:
                        # As the IHH number format for 20t brake vans is not known, choose a random number.
                        number.text = '####B' + str(random.randint(953676, 954520)) + '#'
                    elif hea:
                        # Swap for a Fastline Simulations HEA wagon in railfreight livery
                        rv_int = int(rv_orig[2:6])
                        if 'eTrue' in loaded.text:
                            load = 'L'  # Replace a loaded wagon
                        else:
                            load = 'E'  # Replace an empty wagon
                        if rv_int < 231:
                            blueprint.text = r'RailVehicles\Freight\HEA\HEA_RF_CL_C_' + load + '.xml'
                            name.text = 'HEA_RF_CL_C_' + load
                        else:
                            blueprint.text = r'RailVehicles\Freight\HEA\HEA_RF_OL_C_' + load + '.xml'
                            name.text = 'HEA_RF_OL_C_' + load
                        HEA_RF_suffixes = ['£####(###', '#$###(###', '##%##(###', '###^##)##', '####&amp;#)##']
                        provider.text = 'FastlineSimulation'
                        product.text = 'HBA HEA Hoppers'
                        number.text = 'HEA' + str(360000 + rv_int) + HEA_RF_suffixes[random.randrange(0, 5)]
                        if not tailmarker == 1:
                            # change to a tail lamp carrying wagon and try to orient it with the lamp outward facing
                            if tail_style == 'Flashing':
                                tail_bp = '_Rb.xml'
                                tail_name = '_Rb'
                            else:
                                tail_bp = '_Ro.xml'
                                tail_name = '_Ro'
                            blueprint.text, name.text = add_taillamp(tailmarker, blueprint.text, tail_bp, name.text,
                                                                     tail_name, flipped, followers)
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                        return True
                    elif mcv or tip:
                        provider.text = this_vehicle[3]
                        product.text = this_vehicle[4]
                        random_variant = str(random.randrange(1, 4))
                        blueprint.text = this_vehicle[5].replace('BR 1', 'BR ' + random_variant)
                        name.text = this_vehicle[6].replace('BR 1', 'BR ' + random_variant)
                        number.text = str(550000 + int(rv_orig[2:6]))
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                        return True
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
                        loco = csv_get_blue47num('Class47_dom', rv_num)
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


def ihh_c14_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_14' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class14'])):
                this_vehicle = vehicle_db['IHH_Class14'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    # Replace with a RSC Class 20 and format number accordingly
                    rv_orig = number.text
                    rv_num = str(random.randint(9500, 9555)) + str(random.randint(5, 8)) + 'K' \
                             + str(random.randint(0, 9)) + str(random.randint(0, 9))
                    nm = re.search('^D([0-9]{4})([0-9][a-zA-Z][0-9]{2})', number.text)
                    if nm:
                        rv_num = str((int(nm.group(1)) % 56) + 9500) + nm.group(2).upper()
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = rv_num
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def ihh_c17_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_17' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class17'])):
                this_vehicle = vehicle_db['IHH_Class17'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    # Replace with a RSC Class 20 and format number accordingly
                    rv_orig = number.text
                    rv_num = str(random.randint(9500, 9555)) + str(random.randint(5, 8)) + 'K' \
                             + str(random.randint(0, 9)) + str(random.randint(0, 9))
                    pretops_nm = re.search('^D([0-9]{4})([0-9])[a-zA-Z][0-9]{2}', number.text)
                    if pretops_nm:
                        # A pre-tops loco - only BR Green disc is available as standard DLC
                        # Take a guess that the loco faces cab forwards and append 'F' (forward) not 'R' (rear)
                        rv_num = rsc20headcodes_62_69[pretops_nm.group(2)] + 'F' + pretops_nm.group(1)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = rv_num
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
                    tops_nm = re.search('^.20#([0-9]{3})', number.text)
                    if tops_nm:
                        if int(tops_nm.group(1)) < 127:
                            rv_num = str(20000 + int(tops_nm.group(1)))
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


def ihh_c27_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_27' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class27'])):
                this_vehicle = vehicle_db['IHH_Class27'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_num = '27036'
                    rv_orig = number.text
                    nm_tops = re.search('^(2[6-7][0-9]{3}).*', number.text)
                    nm_pretops = re.search('^D([0-9]{4})([0-9][A-Z][0-9]{2})', number.text)
                    if nm_tops:
                        rv_num = nm_tops.group(1)
                    if nm_pretops:
                        rv_num = nm_pretops.group(1) + nm_pretops.group(2)
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
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                     this_vehicle[7].replace('\\', '/')), rv_tops, '([0-9]{9})(.*)')
                            rv_num = ap_num[0:9] + '2121'
                    else:
                        tops_headcode = re.search('^(40[0-9]{3})(....).*', number.text)
                        if tops_headcode:
                            rv_tops = '11111' + tops_headcode.group(1)
                            ap_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                     this_vehicle[7].replace('\\', '/')), rv_tops, '([0-9]{10})(.*)')
                            hc_search = re.search('([0-9][A-Z][0-9]{2})', tops_headcode.group(2))
                            if hc_search:
                                rv_num = '110' + ap_num[3:10] + hc_search.group(0)
                            else:
                                rv_num = ap_num
                    pretops_num = re.search('D([0-9]{3})([0-9][A-Z][0-9]{2})$', number.text, flags=re.IGNORECASE)
                    pretops_disc = re.search('disc', this_vehicle[2], flags=re.IGNORECASE)
                    if pretops_disc and pretops_num:
                        rv_dnum = '0' + pretops_num.group(1)
                        headcode = ap40headcodes_62_69[pretops_num.group(2)[0:1]]
                        ap_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_dnum, '([0-9]{4})(.*)')
                        hy = re.search('panel', this_vehicle[2], flags=re.IGNORECASE)
                        if hy:
                            # Number the loco as a Half Yellow front class 40
                            rv_num = '1' + ap_num[1:4] + headcode
                        else:
                            # Number the loco as a Full Green front class 40
                            rv_num = '0' + ap_num[1:4] + headcode
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


def ihh_c56_replace(provider, product, blueprint, name, number):
    if 'IHH' in provider.text:
        if 'Class_56' in product.text:
            for i in range(0, len(vehicle_db['IHH_Class56'])):
                this_vehicle = vehicle_db['IHH_Class56'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    number.text = number.text[-5:]
                    rv_list.append(number.text)
                    rv_pairs.append([number.text, number.text])
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
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                     this_vehicle[7].replace('\\', '/')), number.text)
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
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_found, '([0-9]{5})(.*)')
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
                        if not 6700 <= int(rv_dnum) <= 6999:
                            if not 6600 <= int(rv_dnum) <= 6608:
                                # If the pre-tops number in the scenario is not valid invent a new one
                                # find the remainder of the non-valid number divided by 300 and add 6700 - the result
                                # is guaranteed to be in valid range 6700 - 6999
                                rv_dnum = str(6700 + int(rv_dnum) % 300)
                        headcode = pretops.group(2)
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_dnum, '([0-9]{5})(.*)')
                        rv_num = rv_num.replace('____', headcode)
                    # Check if the loco has a tops number
                    tops = re.search('(37[0-9]{3})(.*)', number.text)
                    if tops:
                        rv_tops = tops.group(1)
                        rv_num = this_vehicle[7]
                        if 'dcsv' in rv_num:
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                     this_vehicle[7].replace('\\', '/')), rv_tops, '([0-9]{5})(.*)')
                    if '_wp' in this_vehicle[2]:
                        rv_num = add_ploughs(rv_num)
                    if this_vehicle[1] == 'WHL' or this_vehicle[1] == 'FortWilliamMallaig':
                        if 'Large' in this_vehicle[2]:
                            # Look for a loco with the 'Westie' logo for the WHL LL replacements
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                     this_vehicle[7].replace('\\', '/')), rv_tops, '(37[0-9]{3})(.*L=1.*)')
                        # Add ploughs and RETB to West Highland locos
                        rv_num = add_retb(rv_num)
                        rv_num = add_ploughs(rv_num)
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
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_dnum, '([0-9]{4})(.*)')
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
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_dnum, '([0-9]{4})(.*)')
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
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_tops, '([0-9]{10})(.*)')
                    tops_disc = re.search('^([0-9])(40[0-9]{3})$', number.text)
                    if tops_disc:
                        rv_tops = '1111' + tops_disc.group(2)
                        headcode = ap40headcodes_69_77[tops_disc.group(1)]
                        ap_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_tops, '([0-9]{9})(.*)')
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
    # Replace Class 56 with AP Enhanced version
    for i in range(0, len(vehicle_db['Class56'])):
        this_vehicle = vehicle_db['Class56'][i]
        bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
        if bp:
            rv_orig = number.text
            rv_num = number.text
            if 'RSC' in provider.text and 'Class56Pack01' in product.text:
                rv_num = cl56rsc_to_apsecdep_or_blanksecdep(number.text)
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
                    rv_orig = rv_num = number.text
                    nm = re.search('(66[0-9]{3})', number.text)
                    if nm:
                        rv_found = nm.group(1)
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_found, '([0-9]{5})(.*)')
                    # Set number
                    if len(rv_num) < 6:
                        rv_num = rv_num + 'x'
                    number.text = rv_num
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
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_found, '([0-9]{5})(.*)')
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
                        Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                             this_vehicle[7].replace('\\', '/')), rv_found, '([0-9]{5})(.*)')
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
                    blueprint.text = this_vehicle[5]
                    # name.text = w_name
                    rv_orig = number.text
                    nm = re.match('(86[0-9]{3}).*', number.text)
                    if nm:
                        # Loco to be replaced is TOPS numbered class 86 with no headcode box
                        rv_found = nm.group(1)
                        rv_num = this_vehicle[7]
                        if 'dcsv' in rv_num:
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                     this_vehicle[7].replace('\\', '/')), rv_found, '([0-9]{5})(.*)')
                        # Set number
                        number.text = str(rv_num)
                        name.text = w_name
                        blueprint.text = w_blueprint
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                        return True
                    nm = re.search('(....)(86[0-9]{3})', number.text)
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
                            if nm.group(2) in c86_TOPS_HC:
                                # Look up the TOPS number in the dictionary of VP vehicle numbers and configurations
                                # and use the matching value as the VP railvehicle number
                                rv_num = c86_TOPS_HC[nm.group(2)]
                            else:
                                # TOPS number not found in dictionary so assign the VP railvehicle number below
                                rv_num = nm.group(2) + 'E31830O00;B=E5;P=old;HL=0'
                            # Set the headcode into the VP number
                            rv_num = rv_num.replace('0O00', nm.group(1))
                            # Fix the dot and blank @ and ~ headcode characters where VP use ? and # instead
                            rv_num = rv_num.replace('@', '?')
                            rv_num = rv_num.replace('~', '#')
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
                                     vehicle_db['Class86'][4][7].replace('\\', '/')), nm.group(2), '([0-9]{5})(.*)')
                            # Set number
                            number.text = str(rv_num)
                            rv_list.append(number.text)
                            rv_pairs.append([rv_orig, number.text])
                            return True
                    return False
    return False


def c87_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Class87'])):
        this_vehicle = vehicle_db['Class87'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    (w_blueprint, w_name) = set_weathering(3, this_vehicle)
                    rv_orig = number.text
                    nm = re.match('.?(87[0-9]{3}).*', number.text)
                    if nm:
                        rv_found = nm.group(1)
                        rv_num = rv_found + this_vehicle[7]
                        provider.text = this_vehicle[3]
                        product.text = this_vehicle[4]
                        number.text = str(rv_num)
                        name.text = w_name
                        blueprint.text = w_blueprint
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


def c150_replace(provider, product, blueprint, name, number):
    if 'Thomson_Oovee' in provider.text:
        if 'Class150Pack01' in product.text:
            for i in range(0, len(vehicle_db['DMU150_set'])):
                this_vehicle = vehicle_db['DMU150_set'][i]
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text
                    nm = re.search('5[2,7]([0-9]{3})[0-9]*', number.text)
                    if nm:
                        number.text = '150' + nm.group(1) + 'a'
                    else:
                        # Unit number of the Oovee 150 is not in standard format - can't replace the vehicle
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
                                destination = get_destination(c158_s9bl_rr, nm.group(1), 'a')
                            elif bool(re.search('FGW', blueprint.text, flags=re.IGNORECASE)):
                                destination = get_destination(c158_s9bl_fgw, nm.group(1), 'a')
                            elif bool(re.search('NR', blueprint.text, flags=re.IGNORECASE)):
                                destination = get_destination(c158_s9bl_nr, nm.group(1), 'a')
                            elif bool(re.search('NTPE', blueprint.text, flags=re.IGNORECASE)):
                                destination = get_destination(c158_s9bl_tpe, nm.group(1), 'a')
                            elif bool(re.search('South|SWT', blueprint.text, flags=re.IGNORECASE)):
                                destination = get_destination(c158_s9bl_swt, nm.group(1), 'a')
                            rv_num = nm.group(2) + destination
                    else:
                        nm = re.search('(.)([0-9]{4}).*', number.text)
                        if nm:
                            if (provider.text == 'DTG' and product.text == 'Class158Pack01' and bool(
                                    re.search('Default', blueprint.text, flags=re.IGNORECASE))) or (
                                    provider.text == 'DTG' and product.text == 'NorthWalesCoast' and bool(
                                re.search('Default', blueprint.text, flags=re.IGNORECASE))):
                                # Arriva Trains Wales liveried stock
                                destination = get_destination(c158_nwc, nm.group(1), 'a')
                            elif provider.text == 'DTG' and product.text == 'FifeCircle' and bool(
                                    re.search('Default', blueprint.text, flags=re.IGNORECASE)):
                                # ScotRail saltire liveried stock
                                destination = get_destination(c158_dtg_fc, nm.group(1), 'a')
                            elif provider.text == 'RSC' and product.text == 'LiverpoolManchester' and bool(
                                    re.search('Default', blueprint.text, flags=re.IGNORECASE)):
                                # Regional Railways liveried stock
                                destination = get_destination(c158_livman_rr, nm.group(1), 'a')
                            rv_num = '15' + nm.group(2) + destination
                        if provider.text == 'RSC' and product.text == 'SettleCarlisle':
                            # Destination blank - Settle-Carlisle units don't support destination displays
                            rv_num = '158' + rv_orig[2:5] + 'a'
                    # It's assumed the scenario being converted will have one DMSLA and one DMSLB blueprint in each set
                    # in the consist. If 2 sets or more sets are joined the driving vehicles will alternate DMSLA /
                    # DMSLB.
                    this_bp = this_vehicle[5]
                    this_name = this_vehicle[6]
                    v_type = re.search(r'\\Class158C?_([A-Z]*)', this_vehicle[5], flags=re.IGNORECASE)
                    if v_type:
                        this_bp = alternate_mu_driving_vehicles(this_bp, v_type.group(1), 'DMSLA', 'DMSLB')
                        this_name = get_ap_name_from_bp(vehicle_db['DMU158_set'], this_bp)
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_bp
                    name.text = this_name
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c170_replace(provider, product, blueprint, name, number):
    global mu_last
    for i in range(0, len(vehicle_db['DMU170_set'])):
        this_vehicle = vehicle_db['DMU170_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = rv_num = number.text
                    destination = 'a'
                    nm = re.search('([a-zA-Z]).....([0-9]{6})', number.text)
                    if nm:
                        if bool(re.search(r'\\AR[2-3]\\|\\NXEAWhite\\|\\OR[2-3]\\', blueprint.text,
                                          flags=re.IGNORECASE)):
                            destination = get_destination(c170_ar23, nm.group(1), 'a')
                        elif bool(re.search(r'\\CH\\', blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c168_chiltern, nm.group(1), 'a')
                        elif bool(re.search(r'\\CT\\|\\CTMML\\|\\XC\\', blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c170_ct_xc, nm.group(1), 'a')
                        elif bool(re.search(r'\\LM\\', blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c170_lm, nm.group(1), 'a')
                        elif bool(re.search('Ex-Anglia_Rev_AP|Ex-ONE_AP|Ex-ONE_Dark_AP', blueprint.text,
                                            flags=re.IGNORECASE)):
                            destination = get_destination(c170_ex_ar_aga_ap, nm.group(1), 'a')
                        elif bool(re.search(r'Scotrail|\\FS\\|\\FSRS|\\FSRT|\\SP\\|\\SPSnow\\|',
                                            blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c170_scotrail, nm.group(1), 'a')
                        elif bool(re.search(r'\\GA\\|\\HT\\|\\NXEA\s[2-3]C\\', blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c170_ga_hull, nm.group(1), 'a')
                        elif bool(re.search(r'\\FTPE\\', blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c170_ftpe, nm.group(1), 'a')
                        elif bool(re.search(r'\\MML\\', blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c170_mml, nm.group(1), 'a')
                        elif bool(re.search(r'\\S171\\', blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c171_southern, nm.group(1), 'a')
                        rv_num = nm.group(2) + destination
                    # It's assumed the scenario being converted will have one DMCL and one DMSL blueprintin each set in
                    # the consist. If 2 sets or more sets are joined the driving vehicles will alternate DMCL / DMSL.
                    this_bp = this_vehicle[5]
                    this_name = this_vehicle[6]
                    v_type = re.search(r'\\Class170_([A-Z]*)', this_vehicle[5], flags=re.IGNORECASE)
                    if v_type:
                        # The SR Saltire Class 170 with DMSL is unlike the others so need a workaround
                        if v_type.group(1).upper() == 'DMCLA':
                            this_bp = alternate_mu_driving_vehicles(this_bp, v_type.group(1), 'DMCLA', 'DMSLB')
                        # The Southern Class 170 also needs a workaround
                        elif re.search(r'\\Southern_AP', this_vehicle[5], flags=re.IGNORECASE):
                            this_bp = alternate_mu_driving_vehicles(this_bp, v_type.group(1), 'DMCL', 'DMSLB')
                        else:
                            this_bp = alternate_mu_driving_vehicles(this_bp, v_type.group(1), 'DMCL', 'DMSL')
                        this_name = c170_bp_name_lookup[this_bp]
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_bp
                    name.text = this_name
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c175_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['DMU175_set'])):
        this_vehicle = vehicle_db['DMU175_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = rv_num = number.text
                    nm = re.search('([0-9]{6})([a-zA-Z])', number.text)
                    if nm:
                        # Check if destination is 'v' - Empty to Depot. If so, change to uppercase 'V' used by AP.
                        # Otherwise, destination is consistent with AP scheme and doesn't need changed.
                        if nm.group(2) == 'v':
                            destination = 'V'
                        else:
                            destination = nm.group(2)
                        rv_num = nm.group(1) + destination
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


def c221_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['DMU220-1_set'])):
        this_vehicle = vehicle_db['DMU220-1_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = rv_num = number.text
                    nm_driven = re.search('^..([0-9]{6})([0-9]{5})$', number.text)
                    nm_coach = re.search('^..([0-9]{5})$', number.text)
                    if nm_driven:
                        # Driving vehicle found.
                        rv_num = nm_driven.group(1) + nm_driven.group(2)
                    elif nm_coach:
                        # Coach found
                        rv_num = '221012' + nm_coach.group(1)
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


def c319_replace(provider, product, blueprint, name, number):
    global mu_last, mso_num
    for i in range(0, len(vehicle_db['EMU319_set'])):
        this_vehicle = vehicle_db['EMU319_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    this_bp = this_vehicle[5]
                    this_name = this_vehicle[6]
                    rv_orig = number.text
                    rv_num = rv_orig
                    set_nm = re.search('(319[0-9]{3}).....([a-zA-Z]?)', number.text)
                    if set_nm:
                        if set_nm.group(2) == '':
                            dest = '#'
                        else:
                            dest = set_nm.group(2)
                        mso_num = set_nm.group(1) + dest
                    v_type = re.search(r'\\Class_319_([A-Z]*)', this_vehicle[5], flags=re.IGNORECASE)
                    if v_type:
                        if v_type.group(1).upper() == 'MSO':
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                     this_vehicle[7].replace('\\', '/')), mso_num[0:6], 'Z([0-9]{6})(.*)')
                            rv_num = c319_dest[mso_num[6:]] + rv_num
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_bp
                    name.text = this_name
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c325_replace(provider, product, blueprint, name, number):
    global mu_last
    for i in range(0, len(vehicle_db['EMU325_set'])):
        this_vehicle = vehicle_db['EMU325_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    # It's assumed the scenario being converted will have one DTVA and one DTVB blueprint in each
                    # set in the consist. If 2 sets or more sets are joined the driving vehicles will alternate
                    # DTVA / DTVB.
                    this_bp = this_vehicle[5]
                    this_name = this_vehicle[6]
                    v_type = re.search(r'\\Class325_([A-Z]*)', this_vehicle[5], flags=re.IGNORECASE)
                    if v_type:
                        this_bp = alternate_mu_driving_vehicles(this_bp, v_type.group(1), 'DTVA', 'DTVB')
                        this_name = get_ap_name_from_bp(vehicle_db['EMU325_set'], this_bp)
                    # Check if we're on DC power
                    dc = re.search('_DC', this_vehicle[5], flags=re.IGNORECASE)
                    if dc:
                        this_dcsv = 'PMV_DC.dcsv'
                    else:
                        this_dcsv = 'PMV.dcsv'
                    rv_orig = number.text
                    nm = re.search('[0-9]{5}(325[0-9]{3})', number.text)
                    if nm:
                        rv_num = nm.group(1)
                    else:
                        rv_num = number.text[0:5]
                        if 68340 <= int(rv_num) <= 68355:
                            rv_num = str(int(rv_num) + 256661)
                            rv_num = dcsv_get_num(
                                Path(railworks_path, 'Assets', 'RSC', 'Class325Pack01', 'RailVehicles', 'Class325',
                                     'RM1_W1_AP', this_dcsv), rv_num, '([0-9]{6})(.*)')
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_bp
                    name.text = this_name
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c350_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['EMU350_set'])):
        this_vehicle = vehicle_db['EMU350_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text  # 3503696014219
                    destination = ''
                    nm = re.search('([0-9]{6}).....(.*)', number.text)
                    if nm:
                        if bool(re.search(r'\\FTPE\\', blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c350_lb_ftpe, nm.group(2), '0')
                            if destination == '0':
                                destination = ''
                            else:
                                destination = ';D=' + destination
                        if product.text == 'CrossCity' and \
                                bool(re.search(r'RailVehicles\\Electric\\Class350\\Default\\Engine\\Class350_',
                                               blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c350_lm_cc, nm.group(2), '0')
                            if destination == '0':
                                destination = ''
                            else:
                                destination = ';D=' + destination
                        if product.text == 'WCML-South' and \
                                bool(re.search(r'RailVehicles\\Electric\\Class350\\Default\\Engine\\Class350_',
                                               blueprint.text, flags=re.IGNORECASE)):
                            destination = get_destination(c350_lm_wcmls, nm.group(2), '0')
                            if destination == '0':
                                destination = ''
                            else:
                                destination = ';D=' + destination
                        rv_num = nm.group(1) + destination
                    else:
                        rv_num = number.text[0:6]
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


def c365_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['EMU365_set'])):
        this_vehicle = vehicle_db['EMU365_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = rv_num = number.text
                    destination = 'a'
                    nm = re.search('([0-9]{6})......([a-zA-Z]?)', number.text)
                    # Check if this is an RSC ECMLS 365 format number
                    if nm:
                        if bool(re.search(r'\\Default\\', blueprint.text, flags=re.IGNORECASE)):
                            # This is for the ECMLS 365 NSE livery
                            destination = get_destination(c365_ecmls_nse, nm.group(2), 'a')
                        rv_num = number.text[0:6] + destination
                    nm = re.search('([a-zA-Z]?)........([0-9]{3})', number.text)
                    # Check if this is an RSC Class365Pack02 format number
                    if nm:
                        if bool(re.search(r'\\CXSE_AP\\', blueprint.text, flags=re.IGNORECASE)):
                            # This is for the ECMLS 365 NSE livery
                            destination = get_destination(c365_apcxse, nm.group(1), 'a')
                        rv_num = '365' + nm.group(2) + destination
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


def c375_replace(provider, product, blueprint, name, number):
    global mu_last
    for i in range(0, len(vehicle_db['EMU375-7_set'])):
        this_vehicle = vehicle_db['EMU375-7_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = rv_num = number.text
                    destination = 'a'
                    nm = re.search('([a-zA-Z]).....([0-9]{6})', number.text)
                    if nm:
                        destination = get_destination(c375_dtg_pack, nm.group(1), 'a')
                        if product.text.upper() == 'LondonGillingham':
                            if bool(re.search(r'\\SN\\', blueprint.text, flags=re.IGNORECASE)):
                                # This is for the London-Gillingham Southern livery
                                destination = get_destination(c377_lg_sn, nm.group(1), 'a')
                        if product.text == 'PortsmouthDirect':
                            if bool(re.search(r'\\SN\\', blueprint.text, flags=re.IGNORECASE)):
                                # This is for the Portsmouth Direct Southern livery
                                destination = get_destination(c377_lg_sn, nm.group(1), 'a')
                        if product.text == 'BrightonMainLine':
                            if bool(re.search(r'\\FCC', blueprint.text, flags=re.IGNORECASE)):
                                # This is for the Brighton Main Line FCC livery
                                destination = get_destination(c377_fcc, nm.group(1), 'a')
                            if bool(re.search(r'\\Southern', blueprint.text, flags=re.IGNORECASE)):
                                # This is for the Brighton Main Line Southern livery
                                destination = get_destination(c377_lg_sn, nm.group(1), 'a')
                            if bool(re.search(r'\\SE-White', blueprint.text, flags=re.IGNORECASE)):
                                # This is for the Brighton Main Line SE White livery
                                destination = get_destination(c377_lb_se, nm.group(1), 'a')
                        if product.text == 'WCML-South':
                            if bool(re.search(r'\\Class377\\Engine\\Class377_[A-Z_]*\.xml', blueprint.text,
                                              flags=re.IGNORECASE)):
                                # This is for the DTG WCML South Southern livery
                                destination = get_destination(c375_southern_wcmls, nm.group(1), 'a')
                    # Check whether DMOSA, DMOSB, MOSL, PTOSL, or TOSL and change number accordingly
                    # It's assumed the scenario being converted will have only one DMOSA vehicle at the front and that
                    # all other DMOS vehicles in the consist will be DMOSB. If that's left as it is then the AP ones
                    # will have headlights and taillights on at the same time and generally go to pot.
                    this_bp = this_vehicle[5]
                    this_name = this_vehicle[6]
                    v_type = re.search('375_([A-Z]*)', this_vehicle[5], flags=re.IGNORECASE)
                    if v_type:
                        rv_num = number.text[6:12] + destination
                        if v_type.group(1).upper() == 'DMOSA':
                            # Test if last Driving vehicle was a DxxA - if so, swap this one for a DxxB
                            if mu_last == 'DMOSA':
                                # Swap for a DxxB
                                this_bp = c375_dmos_lookup[this_bp]
                                # Remember that the last driven vehicle in this consist was a DxxB.
                                mu_last = 'DMOSB'
                            else:
                                # Leave the DxxA as it is. Remember that the last driving vehicle processed in this
                                # consist was a DxxA.
                                mu_last = 'DMOSA'
                        elif v_type.group(1).upper() == 'DMOSB':
                            # Test if last Driving vehicle was a DxxB - if so, swap this one for a DxxA
                            if mu_last == 'DMOSB':
                                # Swap for a DxxA
                                this_bp = c375_dmos_lookup[this_bp]
                                # Remember that the last cab vehicle in this consist was a DxxA.
                                mu_last = 'DMOSA'
                            else:
                                # Leave the DxxB as it is. Remember that the last driving vehicle processed in this
                                # consist was a DxxB.
                                mu_last = 'DMOSB'
                        this_name = get_ap_name_from_bp(vehicle_db['EMU375-7_set'], this_bp)
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_bp
                    name.text = this_name
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c450_replace(provider, product, blueprint, name, number):
    global mu_last
    for i in range(0, len(vehicle_db['EMU450_set'])):
        this_vehicle = vehicle_db['EMU450_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text
                    rv_num = number.text[0:6]
                    if provider.text == 'DTG' and product.text == 'PortsmouthDirect':
                        # AP Class 450 uses same destination codes as PortsmouthDirect C450s
                        nm = re.search('([0-9]{6}).....([0-9]{0,2})', number.text)
                        if nm:
                            if len(nm.group(2)) > 0:
                                destination = ';D=' + str(int(nm.group(2)))
                                rv_num = nm.group(1) + destination
                    else:
                        nm = re.search('([0-9]{6}).....([A-Z]{0,1})', number.text)
                        if nm:
                            if len(nm.group(2)) > 0:
                                # London-Brighton and Guildford C450s destinations need translation via dictionary
                                destination = get_destination(c450_gu_swt, nm.group(2), '0')
                                if destination == '0':
                                    destination = ''
                                else:
                                    destination = ';D=' + destination
                                rv_num = nm.group(1) + destination
                    # It's assumed the scenario being converted will have one DMC1 and one DMC2 in each set in the
                    # consist. If 2 sets or more sets are joined the driving vehicles will alternate DMC1 / DMC2.
                    this_bp = this_vehicle[5]
                    this_name = this_vehicle[6]
                    v_type = re.search(r'\\(444|450)_([A-Z]*)', this_vehicle[5], flags=re.IGNORECASE)
                    if v_type:
                        this_bp = alternate_mu_driving_vehicles(this_bp, v_type.group(2), 'DMOS', 'DMOSB')
                        this_name = get_ap_name_from_bp(vehicle_db['EMU450_set'], this_bp)
                    if not bool(re.search('DMC1', this_name, flags=re.IGNORECASE)):
                        # Any vehicle other than a DMOS/DMC1 gets a placeholder number
                        rv_num = v_type.group(2) + number.text[0:6]
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_bp
                    name.text = this_name
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c456_replace(provider, product, blueprint, name, number):
    global mu_last
    for i in range(0, len(vehicle_db['EMU456_set'])):
        this_vehicle = vehicle_db['EMU456_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = number.text
                    rv_num = number.text[0:11]
                    rv_dest = number.text[11]
                    if bool(re.search(r'\\NetworkSE\\', blueprint.text, flags=re.IGNORECASE)):
                        rv_num = c456_nse[rv_dest] + rv_num
                    elif bool(re.search(r'\\Default\\', blueprint.text, flags=re.IGNORECASE)):
                        rv_num = c456_southern[rv_dest] + rv_num
                    else:
                        # At the moment only the Default (Southern) 456 and the NSE 456 are configured
                        return False
                    # It's assumed the scenario being converted will have one DMC1 and one DMC2 in each set in the
                    # consist. If 2 sets or more sets are joined the driving vehicles will alternate DMC1 / DMC2.
                    this_bp = this_vehicle[5]
                    this_name = this_vehicle[6]
                    v_type = re.search(r'\\Class_456_([A-Z]*)', this_vehicle[5], flags=re.IGNORECASE)
                    if v_type:
                        this_bp = alternate_mu_driving_vehicles(this_bp, v_type.group(1), 'DMSO', 'DTSO')
                        this_name = get_ap_name_from_bp(vehicle_db['EMU456_set'], this_bp)
                    # Swap vehicle and set number / destination (where possible)
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_bp
                    name.text = this_name
                    number.text = str(rv_num)
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def c465_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['EMU465_set'])):
        this_vehicle = vehicle_db['EMU465_set'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    rv_orig = rv_num = number.text
                    destination = 'a'
                    nm = re.search('([a-zA-Z]?)........([0-9]{3})', number.text)
                    if nm:
                        destination = get_destination(c465_se, nm.group(1), 'a')
                        rv_num = '465' + nm.group(2) + destination
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
    global mu_last
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
    progress_win.bring_to_front()
    progress_win.force_focus()
    progress_bar = progress_win.find_element('progress')
    consist_nr = 0
    for citem in consists:
        service = citem.find('Driver/cDriver/ServiceName/Localisation-cUserLocalisedString/English')
        if service is None:
            service = 'Loose consist'
            driven = False
        else:
            service = service.text
            driven = True
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
            consist_item_nr = 0
            for coentity in rvehicles.findall('cOwnedEntity'):
                consist_item_nr = consist_item_nr + 1
                consist_items_total = len(rvehicles.findall('cOwnedEntity'))
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
                flipped = coentity.find('Component/*/Flipped')
                followers = coentity.find('Component/*/Followers')
                # Set tailmarker - 0 if vehicle is first in consist, 2 if last, 1 if somewhere in between
                # We need this if the vehicles with tail lights come in a separate blueprint (e.g. AP HAA).
                # A vehicle will not get a tail light added if it is in a loose consist.
                tailmarker = 1
                if driven == True:
                    if consist_item_nr == 1:
                        tailmarker = 0
                    elif consist_item_nr == consist_items_total:
                        tailmarker = 2
                input_vehicle_list.append(
                    [str(consist_nr), provider.text, product.text, blueprint.text, name.text, number.text, loaded.text,
                     service, playerdriven])
                vehicle_replacer(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker)
                output_vehicle_list.append(
                    [str(consist_nr), provider.text, product.text, blueprint.text, name.text, number.text, loaded.text,
                     service, playerdriven])
            mu_last = 'none'
            mso_num = ''
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


def vehicle_replacer(provider, product, blueprint, name, number, loaded, flipped, followers, tailmarker):
    # Check the rail vehicle found by the XML parser against each of the enabled substitutions.
    # A soon as a replacement is made, return to the XML parser and search for the next vehicle.
    if values['Replace_Mk1']:
        mk1_replace(provider, product, blueprint, name, number)
    if values['Replace_Mk2ac'] and mk2ac_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_Mk2df'] and mk2df_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_Mk3ab'] and mk3ab_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_FSA'] and fsafta_replace(provider, product, blueprint, name, number, loaded, flipped, followers,
                                                tailmarker):
        return True
    if values['Replace_HAA'] and haa_replace(provider, product, blueprint, name, number, loaded, flipped, followers,
                                             tailmarker):
        return True
    if values['Replace_HHA'] and hha_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_HTO'] and coal21_t_hto_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_HTV'] and coal21_t_htv_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_TTA'] and tta_replace(provider, product, blueprint, name, number, loaded):
        return True
    if values['Replace_VDA'] and vda_replace(provider, product, blueprint, name, number, loaded, flipped, followers,
                                             tailmarker):
        return True
    if values['Replace_IHH'] and ihh_replace(provider, product, blueprint, name, number, loaded, flipped, followers,
                                             tailmarker):
        return True
    if values['Replace_User'] and user_replace(provider, product, blueprint, name):
        return True
    if values['Replace_HST'] and hst_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C91'] and c91_replace(provider, product, blueprint, name):
        return True
    if values['Replace_C101'] and c101_replace(provider, product, blueprint, name):
        return True
    if values['Replace_C150'] and c150_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C156'] and c156_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C158'] and c158_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C170'] and c170_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C175'] and c175_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C221'] and c221_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C319'] and c319_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C325'] and c325_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C350'] and c350_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C365'] and c365_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C375'] and c375_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C450'] and c450_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C456'] and c456_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_C465'] and c465_replace(provider, product, blueprint, name, number):
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
    if values['Replace_C87'] and c87_replace(provider, product, blueprint, name, number):
        return True
    return True


def fix_short_tags(xml_string):
    # This clumsy fix is necessary because sometimes TS requires short xml empty tags and sometimes long ones.
    # The following substitutions should take care of the important exceptions to the long tag default.
    xml_string = re.sub(r'(<cEngineSimContainer[^>]*)></cEngineSimContainer>', r'\1/>', xml_string,
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
td.input,th.input {
    color: #2222bb;
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
    if config.get('defaults', 'save_report') == report_opts[1]:
        # User wants a basic report just containing details of the vehicles in the processed scenario
        htmrv = "<h1>Rail vehicle list</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n" \
                "    <tr style=\"text-align: right;\">\n      <th>Consist</th>\n      <th>Provider</th>\n" \
                "      <th>Product</th>\n      <th>Blueprint</th>\n      <th>Name</th>\n      <th>Number</th>\n" \
                "      <th>Loaded</th>\n    </tr>\n  </thead>\n  <tbody>\n"
    else:
        # Full report wanted containing details of the vehicles in the processed scenario as well as original vehicles
        htmrv = "<h1>Rail vehicle swap list</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n" \
                "    <tr style=\"text-align: right;\">\n      <th>Consist</th>\n" \
                "      <th class=\"input\">Original Provider</th>\n      <th class=\"input\">Original Product</th>\n" \
                "      <th class=\"input\">Original Blueprint</th>\n      <th class=\"input\">Original Name</th>\n" \
                "      <th class=\"input\">Original Number</th>\n      <th class=\"input\">Loaded</th>\n" \
                "      <th>New Provider</th>\n      <th>New Product</th>\n      <th>New Blueprint</th>\n" \
                "      <th>New Name</th>\n      <th>New Number</th>\n      <th>Loaded</th>\n    </tr>\n  </thead>\n" \
                "  <tbody>\n"
    unique_assets = []
    last_cons = -1
    row_no = 0
    for row in output_vehicle_list:
        if row[1:3] not in unique_assets:
            unique_assets.append(row[1:3])
        if int(row[0]) > last_cons:
            # start of a new consist - count how many vehicles are in this consist
            rowspan = (list(zip(*output_vehicle_list))[0]).count(row[0])
        else:
            rowspan = 0
        col_htm = ''
        row[3] = row[3].replace('.xml', '.bin')
        if Path(railworks_path, 'Assets', row[1], row[2], row[3].replace('\\', '/')).is_file():
            tdstyle = ''
        else:
            tdstyle = ' class="missing"'
        cname = '<i>' + str(row[7]) + '</i>'
        if row[8] is True:
            # Consist is driven by the player - make the name bold and append (Player driven)
            cname = '<b>' + cname + '</b> (Player driven)'
        if rowspan > 0:
            col_htm = col_htm + '      <td rowspan=' + str(rowspan) + '>' + cname + '</td>\n'
        if config.get('defaults', 'save_report') == report_opts[2]:
            # User wants a full report so add columns with details of original vehicles to right hand side of table
            in_row = input_vehicle_list[row_no]
            in_row[3] = in_row[3].replace('.xml', '.bin')
            for col in in_row[1:7]:
                col_htm = col_htm + '      <td class="input">' + col + '</td>\n'
        for col in row[1:7]:
            col_htm = col_htm + '      <td' + tdstyle + '>' + col + '</td>\n'
        if (int(row[0]) % 2) == 0:
            htmrv = htmrv + '    <tr>\n' + col_htm + '    </tr>\n'
        else:
            htmrv = htmrv + '    <tr class=\"shaded_row\">\n' + col_htm + '    </tr>\n'
        last_cons = int(row[0])
        row_no = row_no + 1
    htmrv = htmrv + '  </tbody>\n</table>\n<h3>' + str(
        len(output_vehicle_list)) + ' vehicles in total in this scenario.</h3>'
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
                                                                                     '    <tr>\n      <th>Description</th>\n      <td>' + str(
            scenarioProps[1]) + '</td>\n    </tr>\n' \
                                '    <tr>\n      <th>Briefing</th>\n      <td>' + str(
            scenarioProps[2]) + '</td>\n    </tr>\n' \
                                '    <tr>\n      <th>Start From</th>\n      <td>' + str(
            scenarioProps[3]) + '</td>\n    </tr>\n' \
                                '    <tr>\n      <th>Route</th>\n      <td>' + str(
            scenarioProps[4]) + '</td>\n    </tr>\n' \
                                '  </table>\n'
    htm = htmhead + htp + htmas + htmrv + '</body>\n</html>\n'
    html_file_path.touch()
    html_file_path.write_text(htm)
    return True


if __name__ == "__main__":
    window = sg.Window('RSSwapTool - Rolling stock swap tool', layout)
    try:
        os.chdir(Path(railworks_path, 'Content', 'Routes'))
    except:
        sg.PopupError(str(Path(railworks_path, 'Content', 'Routes')) + ' not found. Please go into Settings and check '
                                                                       'the path to the RailWorks directory is correct')
    while True:
        event, values = window.read()
        if event == 'Exit' or event == sg.WIN_CLOSED:
            break
        elif event == 'About':
            sg.Popup('About RSSwapTool',
                     'Tool for swapping rolling stock in Train Simulator (Dovetail Games) scenarios',
                     'Version ' + version_number + ' / ' + version_date,
                     'Copyright 2023 JR McKenzie (jrmknz@yahoo.co.uk)', 'https://github.com/jrmckenzie/RSSwapTool',
                     'This program is free software: you can redistribute it and / or modify '
                     'it under the terms of the GNU General Public License as published by '
                     'the Free Software Foundation, either version 3 of the License, or '
                     '(at your option) any later version.',
                     'This program is distributed in the hope that it will be useful, '
                     'but WITHOUT ANY WARRANTY; without even the implied warranty of '
                     'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the '
                     'GNU General Public License for more details.',
                     'You should have received a copy of the GNU General Public License '
                     'along with this program.  If not, see https://www.gnu.org/licenses/.')
        elif event == 'Settings':
            if not config.has_section('defaults'):
                config.add_section('defaults')
            config.set('defaults', 'replace_mk1', str(values['Replace_Mk1']))
            config.set('defaults', 'replace_mk2ac', str(values['Replace_Mk2ac']))
            config.set('defaults', 'replace_mk2df', str(values['Replace_Mk2df']))
            config.set('defaults', 'replace_mk3ab', str(values['Replace_Mk3ab']))
            config.set('defaults', 'replace_fsa', str(values['Replace_FSA']))
            config.set('defaults', 'replace_haa', str(values['Replace_HAA']))
            config.set('defaults', 'replace_hha', str(values['Replace_HHA']))
            config.set('defaults', 'replace_hto', str(values['Replace_HTO']))
            config.set('defaults', 'replace_htv', str(values['Replace_HTV']))
            config.set('defaults', 'replace_tta', str(values['Replace_TTA']))
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
            config.set('defaults', 'replace_c87', str(values['Replace_C87']))
            config.set('defaults', 'replace_hst', str(values['Replace_HST']))
            config.set('defaults', 'replace_c91', str(values['Replace_C91']))
            config.set('defaults', 'replace_c101', str(values['Replace_C101']))
            config.set('defaults', 'replace_c150', str(values['Replace_C150']))
            config.set('defaults', 'replace_c156', str(values['Replace_C156']))
            config.set('defaults', 'replace_c158', str(values['Replace_C158']))
            config.set('defaults', 'replace_c170', str(values['Replace_C170']))
            config.set('defaults', 'replace_c175', str(values['Replace_C175']))
            config.set('defaults', 'replace_c221', str(values['Replace_C221']))
            config.set('defaults', 'replace_c319', str(values['Replace_C319']))
            config.set('defaults', 'replace_c325', str(values['Replace_C325']))
            config.set('defaults', 'replace_c350', str(values['Replace_C350']))
            config.set('defaults', 'replace_c365', str(values['Replace_C365']))
            config.set('defaults', 'replace_c375', str(values['Replace_C375']))
            config.set('defaults', 'replace_c450', str(values['Replace_C450']))
            config.set('defaults', 'replace_c456', str(values['Replace_C456']))
            config.set('defaults', 'replace_c465', str(values['Replace_C465']))
            with open(path_to_config, 'w') as configfile:
                config.write(configfile)
                configfile.close()
            c86_hc = config.get('defaults', 'c86_hc', fallback=c86_opts[0])
            c56_rf = config.get('defaults', 'c56_rf', fallback=c56_opts[0])
            fsafta_variant = config.get('defaults', 'fsafta_variant', fallback=fsafta_opts[0])
            fsafta_hc = config.get('defaults', 'fsafta_hc', fallback=fsafta_cube[0])
            mgr_variant = config.get('defaults', 'mgr_variant', fallback=mgr_types[0])
            mgr_livery = config.get('defaults', 'mgr_livery', fallback=mgr_liveries[0])
            vda_livery = config.get('defaults', 'vda_livery', fallback=vda_liveries[0])
            vda_whiteroof_probability = config.get('defaults', 'vda_whiteroof_probability',
                                                   fallback=vda_whiteroof_probabilities[0])
            vda_dirty_probability = config.get('defaults', 'vda_dirty_probability',
                                               fallback=dirty_probabilities[8])
            htx_dirty_probability = config.get('defaults', 'htx_dirty_probability',
                                               fallback=dirty_probabilities[8])
            htx_era = config.get('defaults', 'htx_era', fallback=htx_eras[1])
            tail_style = config.get('defaults', 'tail_style', fallback=tail_opts[0])
            save_report = config.get('defaults', 'save_report', fallback=report_opts[0])
            # The settings button has been pressed, so allow the user to change the RailWorks folder setting
            loclayout = [
                [sg.Text('Settings', font='Helvetica 14')],
                [sg.Text('Path to RailWorks folder:'),
                 sg.Input(default_text=str(railworks_path), key='RWloc', readonly=True),
                 sg.FolderBrowse(key='RWloc')],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text(
                    'Swap behaviour for RSC Class 86 Pack locos with headcode blinds:', size=(72, 0))],
                [sg.Combo(c86_opts, auto_size_text=True, default_value=c86_hc, key='c86_hc', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text(
                    'Railfreight Sectors Class 56: if depot plaque / sectors logo not present in AP Enhancement Pack:',
                    size=(72, 0))],
                [sg.Combo(c56_opts, auto_size_text=True, default_value=c56_rf, key='c56_rf', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text(
                    'If FSA/FTA wagons are found, replace them with the following types:',
                    size=(72, 0))],
                [sg.Combo(fsafta_opts, auto_size_text=True, default_value=fsafta_variant, key='fsafta_variant',
                          readonly=True),
                 sg.Combo(fsafta_cube, auto_size_text=True, default_value=fsafta_hc, key='fsafta_hc', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text(
                    'If HTO/HTV wagons are found, use data panels: pre-TOPS only, TOPS only, or mixed:',
                    size=(72, 0))],
                [sg.Combo(htx_eras, auto_size_text=True, default_value=htx_era, key='htx_era', readonly=True),
                 sg.Text('% chance wagon is dirty:'), sg.Combo(dirty_probabilities, auto_size_text=True,
                                                               default_value=htx_dirty_probability,
                                                               key='htx_dirty_probability', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text(
                    'If HAA wagons are found, replace them with the following type(s):',
                    size=(72, 0))],
                [sg.Combo(mgr_types, auto_size_text=True, default_value=mgr_variant, key='mgr_variant', readonly=True),
                 sg.Text('Livery:'), sg.Combo(mgr_liveries, auto_size_text=True, default_value=mgr_livery,
                                              key='mgr_livery', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text(
                    'If VDA wagons are found, replace them with the following livery, and choose the percentage '
                    'probability that a swapped wagon will have a white painted roof:',
                    size=(72, 0))],
                [sg.Text('Livery:'), sg.Combo(vda_liveries, auto_size_text=True, default_value=vda_livery,
                                              key='vda_livery', readonly=True), sg.Text('% chance roof is white:'),
                 sg.Combo(vda_whiteroof_probabilities, auto_size_text=True,
                          default_value=vda_whiteroof_probability, key='vda_whiteroof_probability', readonly=True),
                 sg.Text('% chance van is dirty:'), sg.Combo(dirty_probabilities, auto_size_text=True,
                                                             default_value=vda_dirty_probability,
                                                             key='vda_dirty_probability', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text('HEA and some VDA wagons can have either steady or flashing (1970s onwards) red tail lamp '
                         'styles. \nIn the event there is a choice, which type would you prefer to see:'),
                 sg.Combo(tail_opts, auto_size_text=True, default_value=tail_style, key='tail_style', readonly=True)],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text("Save a report of vehicles in the scenario"),
                 sg.Combo(report_opts, auto_size_text=True, default_value=save_report, key='save_report', readonly=True,
                          tooltip='You may choose to save a report listing all the rail vehicles (and their numbers)'
                                  ' in the scenario, in .html format, alongside the scenario output file. Optionally,'
                                  ' the report can include details of both the original vehicle and the vehicle after'
                                  ' processing by this application.')],
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
                    config.set('defaults', 'fsafta_variant', str(lvalues['fsafta_variant']))
                    config.set('defaults', 'fsafta_hc', str(lvalues['fsafta_hc']))
                    config.set('defaults', 'mgr_variant', str(lvalues['mgr_variant']))
                    config.set('defaults', 'mgr_livery', str(lvalues['mgr_livery']))
                    config.set('defaults', 'vda_livery', str(lvalues['vda_livery']))
                    config.set('defaults', 'vda_whiteroof_probability', str(lvalues['vda_whiteroof_probability']))
                    config.set('defaults', 'vda_dirty_probability', str(lvalues['vda_dirty_probability']))
                    config.set('defaults', 'htx_dirty_probability', str(lvalues['htx_dirty_probability']))
                    config.set('defaults', 'htx_era', str(lvalues['htx_era']))
                    config.set('defaults', 'tail_style', str(lvalues['tail_style']))
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
            config.set('defaults', 'replace_mk3ab', str(values['Replace_Mk3ab']))
            config.set('defaults', 'replace_fsa', str(values['Replace_FSA']))
            config.set('defaults', 'replace_haa', str(values['Replace_HAA']))
            config.set('defaults', 'replace_hha', str(values['Replace_HHA']))
            config.set('defaults', 'replace_hto', str(values['Replace_HTO']))
            config.set('defaults', 'replace_htv', str(values['Replace_HTV']))
            config.set('defaults', 'replace_tta', str(values['Replace_TTA']))
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
            config.set('defaults', 'replace_c87', str(values['Replace_C87']))
            config.set('defaults', 'replace_hst', str(values['Replace_HST']))
            config.set('defaults', 'replace_c91', str(values['Replace_C91']))
            config.set('defaults', 'replace_c101', str(values['Replace_C101']))
            config.set('defaults', 'replace_c150', str(values['Replace_C150']))
            config.set('defaults', 'replace_c156', str(values['Replace_C156']))
            config.set('defaults', 'replace_c158', str(values['Replace_C158']))
            config.set('defaults', 'replace_c170', str(values['Replace_C170']))
            config.set('defaults', 'replace_c175', str(values['Replace_C175']))
            config.set('defaults', 'replace_c221', str(values['Replace_C221']))
            config.set('defaults', 'replace_c319', str(values['Replace_C319']))
            config.set('defaults', 'replace_c325', str(values['Replace_C325']))
            config.set('defaults', 'replace_c350', str(values['Replace_C350']))
            config.set('defaults', 'replace_c365', str(values['Replace_C365']))
            config.set('defaults', 'replace_c375', str(values['Replace_C375']))
            config.set('defaults', 'replace_c450', str(values['Replace_C450']))
            config.set('defaults', 'replace_c456', str(values['Replace_C456']))
            config.set('defaults', 'replace_c465', str(values['Replace_C465']))
            with open(path_to_config, 'w') as configfile:
                config.write(configfile)
                configfile.close()
            if len(values['Scenario_xml']) < 1:
                sg.popup('No scenario selected!')
            else:
                scenarioPath = Path(values['Scenario_xml'])
                scenarioProperties = scenarioPath.parent / 'ScenarioProperties.xml'
                os.chdir(scenarioPath.parent)
                outPathStem = scenarioPath.parent / Path(str(scenarioPath.stem) + '-' + time.strftime('%Y%m%d-%H%M%S'))
                inFile = scenarioPath
                cmd = railworks_path / Path('serz.exe')
                serz_output = ''
                output_message = ''
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
                    inFileW = str(PureWindowsPath(inFile))
                    scenarioPathW = str(PureWindowsPath(scenarioPath))
                    if platform.system() == 'Windows':
                        # Operating system is Microsoft Windows
                        p1 = subprocess.Popen([str(cmd), scenarioPathW, '/xml:' + inFileW], stdout=subprocess.PIPE)
                    elif platform.system() == 'Linux' and platform.release()[-5:-1] == 'WSL2':
                        # Operating system is Windows Subsystem Linux (WSL2)
                        # Linux-style pathnames can be converted to windows style with drive letters
                        inFileW = inFileW[5] + ':' + inFileW[6:]
                        scenarioPathW = scenarioPathW[5] + ':' + scenarioPathW[6:]
                        p1 = subprocess.Popen([str(cmd), scenarioPathW, '/xml:' + inFileW], stdout=subprocess.PIPE)
                    else:
                        # Operating system has wine to run serz.eze
                        try:
                            wine_executable
                        except NameError:
                            wine_executable = '/usr/bin/wine'
                        p1 = subprocess.Popen([wine_executable, str(cmd), scenarioPathW, '/xml:' +
                                               inFileW], stdout=subprocess.PIPE)
                    p1.wait()
                    # Uncomment the line below to see the output of the serz.exe command
                    # serz_output = 'serz.exe ' + p1.communicate()[0].decode('ascii')
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
                xmlFile.write_text(xmlString, encoding='utf-8')
                output_message = 'Scenario converted.\n'
                html_report_status_text = ''
                if str(scenarioPath.suffix) == '.bin':
                    # Run the serz.exe command again to generate the output scenario .bin file
                    binFile = scenarioPath.parent / Path(str(scenarioPath.stem) + '.bin')
                    binFileW = str(PureWindowsPath(binFile))
                    xmlFileW = str(PureWindowsPath(xmlFile))
                    if platform.system() == 'Windows':
                        # Operating system is Microsoft Windows
                        p2 = subprocess.Popen([str(cmd), xmlFileW, '/bin:' + binFileW], stdout=subprocess.PIPE)
                    elif platform.system() == 'Linux' and platform.release()[-5:-1] == 'WSL2':
                        # Operating system is Windows Subsystem Linux (WSL2)
                        # Linux-style pathnames can be converted to windows style with drive letters
                        binFileW = binFileW[5] + ':' + binFileW[6:]
                        xmlFileW = xmlFileW[5] + ':' + xmlFileW[6:]
                        p2 = subprocess.Popen([str(cmd), xmlFileW, '/bin:' + binFileW], stdout=subprocess.PIPE)
                    else:
                        # Operating system has wine to run serz.eze
                        try:
                            wine_executable
                        except NameError:
                            wine_executable = '/usr/bin/wine'
                        p2 = subprocess.Popen([wine_executable, str(cmd), xmlFileW, '/bin:' +
                                               binFileW], stdout=subprocess.PIPE)
                    p2.wait()
                    inFile.unlink()
                    # Uncomment the following line to see the output of the serz.exe command
                    # output_message = serz_output + '\nserz.exe ' + p2.communicate()[0].decode('ascii')
                output_message = output_message + \
                                 '\nOriginal scenario backup located in ' + str(outPathStem) + str(scenarioPath.suffix)
                if not config.get('defaults', 'save_report') == report_opts[0]:
                    # The user wants a report to be generated
                    scenario_properties = parse_properties_xml(scenarioPath.parent)
                    html_report_file = scenarioPath.parent / Path(str(scenarioPath.stem) + '-railvehicle_report.html')
                    convert_vlist_to_html_table(html_report_file, scenario_properties)
                    html_report_status_text = 'Report listing all rail vehicles located in ' + str(html_report_file)
                    browser = sg.popup_yes_no(output_message, html_report_status_text,
                                              'Do you want to open the report in your web browser now?')
                    if browser == 'Yes':
                        webbrowser.open(html_report_file.as_uri())
                else:
                    sg.popup(output_message)
                # re-initialise all vehicle lists
                rv_list = []
                rv_pairs = []
                output_vehicle_list = []
                input_vehicle_list = []
