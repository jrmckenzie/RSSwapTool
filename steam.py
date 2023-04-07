#     RSSwapTool - A script to swap in up to date or enhanced rolling stock
#     for older versions of stock in Train Simulator scenarios.
#     Copyright (C) 2021 James McKenzie jrmknz@yahoo.co.uk
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
    c456_nse, c456_southern, c319_dest, c350_lm_wcmls, c375_southern_wcmls

# If you want to run this script on Linux you must enter the path to the wine executable. You need wine in order to
# run the serz.exe utility which converts between .bin and .xml scenario files.
# If you're not running this script on Linux this line should be left as the default.
wine_executable = '/usr/bin/wine'

# Initialise the script, set the look and feel and get the configuration
version_number = '1.0.9'
version_date = '8 April 2023'

# Initialise the script, set the look and feel and get the configuration
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
#report_opts = ['Don\'t save a report', 'Save a report']
report_opts = ['Don\'t save a report', 'Save details of processed scenario only',
               'Save details of original and processed scenarios']
sg.LOOK_AND_FEEL_TABLE['Railish'] = {'BACKGROUND': '#21301b',
                                     'TEXT': '#FFFFFF',
                                     'INPUT': '#FFFFFF',
                                     'TEXT_INPUT': '#000000',
                                     'SCROLL': '#944f40',
                                     'BUTTON': ('#FFFFFF', '#000804'),
                                     'PROGRESS': ('#316948', '#003317'),
                                     'BORDER': 2, 'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 2, }
sg.theme('Railish')
config = configparser.ConfigParser()
script_path = Path(os.path.abspath(os.path.dirname(sys.argv[0])))
path_to_config = script_path / 'config_steam.ini'
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
    config.set('defaults', 'replace_black5', 'True')
    config.set('defaults', 'replace_lms8f', 'True')
    config.set('defaults', 'replace_r04', 'True')
    config.set('defaults', 'replace_gwr57xx', 'True')
    config.set('defaults', 'replace_lms5xp', 'True')
    config.set('defaults', 'replace_bulleidlp', 'True')
    config.set('defaults', 'replace_bulleidrlp', 'True')
    config.set('defaults', 'replace_dtmaunsell', 'True')
    config.set('defaults', 'replace_srn15', 'True')
    config.set('defaults', 'save_report', report_opts[0])
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
        with open((Path(os.path.realpath(__file__)).parent / csv_filename), 'r') as csv_file:
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
user_db_path = Path(os.path.realpath(__file__)).parent / 'tables/User.csv'
if not user_db_path.is_file():
    head = 'Label,Provider,Product,Blueprint,ReplaceProvider,ReplaceProduct,ReplaceBlueprint,ReplaceName,NumbersDcsv\n'
    user_db_path.touch()
    user_db_path.write_text(head)
vehicle_db = import_data_from_csv('tables/Steam.csv')
user_db = import_data_from_csv('tables/User.csv')

# Set the layout of the GUI
left_column = [
    [sg.Text('RSSwapTool-steam', font='Helvetica 16')],
    [sg.Text('© 2023 JR McKenzie', font='Helvetica 7')],
    [sg.FileBrowse('Select scenario file to process', key='Scenario_xml', tooltip='Locate the scenario .bin or .xml '
                                                                                  'file you wish to process')],
    [sg.Text('Tick the boxes below to choose the\nsubstitutions you would like to make.')],
]
mid_column = [
    [sg.Checkbox('Replace LMS Black 5', default=get_my_config_boolean('defaults', 'replace_black5'), enable_events=True,
                 tooltip='Tick to enable replacing of Black 5 with BMG LMS 5MT',
                 key='Replace_Black5')],
    [sg.Checkbox('Replace LMS 8F', default=get_my_config_boolean('defaults', 'replace_lms8f'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of 8F with BMG LMS 8F', key='Replace_LMS8F')],
    [sg.Checkbox('Replace Robinson O4', default=get_my_config_boolean('defaults', 'replace_r04'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of O4 with CW GCR 8K', key='Replace_R04')],
    [sg.Checkbox('Replace GWR 57xx', default=get_my_config_boolean('defaults', 'replace_gwr57xx'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of GWR Pannier Tanks with Victory Works Pannier',
                 key='Replace_GWR57xx')],
    [sg.Checkbox('Replace LMS 5XP', default=get_my_config_boolean('defaults', 'replace_lms5xp'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Jubilee with BMG Jubilee', key='Replace_LMS5XP')],
    [sg.Checkbox('Replace Bulleid LP', default=get_my_config_boolean('defaults', 'replace_bulleidlp'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Bulleid LP with BMG Bulleid LP', key='Replace_BulleidLP')],
    [sg.Checkbox('Replace Bulleid RLP', default=get_my_config_boolean('defaults', 'replace_bulleidrlp'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of Bulleid RLP with BMG Bulleid RLP', key='Replace_BulleidRLP')],
    [sg.Checkbox('Replace Maunsell coaches', default=get_my_config_boolean('defaults', 'replace_dtmaunsell'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of DT Maunsell coaches with MatrixTrains', key='Replace_DTMaunsell')],
    [sg.Checkbox('Replace N15 777 Sir Lamiel', default=get_my_config_boolean('defaults', 'replace_srn15'),
                 enable_events=True,
                 tooltip='Tick to enable replacing of RSC N15 with Caledonia Works N15',
                 key='Replace_SRN15')],
]
right_column = [
        [sg.Button('Replace!')],
        [sg.Button('Settings')],
        [sg.Button('About')],
        [sg.Button('Exit')]
]

# Set the layout of the application window
layout = [
    [
        sg.Column(left_column),
        sg.VSeperator(color='#944f40'),
        sg.Column(mid_column),
        sg.VSeperator(color='#944f40'),
        sg.Column(right_column),
    ]
]


def dcsv_get_num(this_dcsv, this_rv, this_re):
    # Try to retrieve the closest match for the loco number from the vehicle number database
    try:
        dcsv_tree = ET.parse(this_dcsv)
    except FileNotFoundError:
        sg.popup('Vehicle number database ' + str(Path(this_dcsv)) + ' not found.',
                 'Check you have all the required products installed, and that you have clicked "Settings" in this '
                 'program and set the location of your RailWorks folder correctly.',
                 'This program will now quit.', title='Error')
        sys.exit('Fatal Error: Vehicle number database ' + str(Path(this_dcsv)) + ' not found.')
    except ET.ParseError:
        sg.popup('Vehicle number database ' + str(Path(this_dcsv)) + ' was found but could not be parsed.',
                 'This program will now quit.', title='Error')
        sys.exit(
            'Fatal Error: Vehicle number database ' + str(Path(this_dcsv)) + ' was found but could not be parsed.')
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


def set_weathering(this_weather_variant, this_vehicle):
    if this_weather_variant == 2:
        weather = 'W' + str(random.randint(1, 2))
        return this_vehicle[5].replace('W2', weather), this_vehicle[6].replace('W2', weather)
    elif this_weather_variant == 3:
        weather = 'W' + str(random.randint(1, 3))
        return this_vehicle[5].replace('W1', weather), this_vehicle[6].replace('W1', weather)
    return False


def get_ap_name_from_bp(this_vehicle_db, this_bp):
    for i in range(0,len(this_vehicle_db)):
        if this_bp.upper() == this_vehicle_db[i][5].upper():
            return this_vehicle_db[i][6]
    return False


def black5_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['Black5'])):
        this_vehicle = vehicle_db['Black5'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    nm = re.search('([A-Z]?)4([0-9]{4})', number.text)
                    if nm:
                        rv_lamp = nm.group(1)
                        if not rv_lamp:
                            rv_lamp = 'B'
                        rv_num = nm.group(2) + '#' + rv_lamp + 'N'
                        # Check if number belongs to a domeless loco
                        if 5224 >= int(nm.group(2)) >= 5000:
                            # Domeless loco
                            product.text = 'Black5Pack03'
                            name.text = this_vehicle[6] + ' Domeless'
                        # Set number
                        number.text = rv_num
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def lms8f_replace(provider, product, blueprint, name, number):
    rv_num = number.text
    nm = re.search('([A-Z])([0-9]{2})#4([0-9]{4})', rv_num)
    if nm:
        rv_num = nm.group(3) + '#NB#5#' + nm.group(2) + nm.group(1) + '0CN'
        result = vehicle_replace(vehicle_db['LMS8F'], provider, product, blueprint, name, number, rv_num)
        return result
    return False


def gcr8k_replace(provider, product, blueprint, name, number):
    rv_num = number.text
    nm = re.search('([A-Z])([0-9]{2})#([0-9]{5})', rv_num)
    if nm:
        rv_num = 'RN=' + nm.group(3) + ';SC=' + nm.group(2) + nm.group(1)
        result = vehicle_replace(vehicle_db['GCR8K'], provider, product, blueprint, name, number, rv_num)
        return result
    return False


def gwr57xx_replace(provider, product, blueprint, name, number):
    rv_num = number.text
    nm = re.search('([0-9]{4})', rv_num)
    if nm:
        rv_num = nm.group(1) + '83A11Y1YY22YNY2Y3Y1YY01NC'
        if product == 'FalmouthBranch':
            rv_num = nm.group(1) + '83F12Y1YY22YNY2Y3Y1YY01NK'
        result = vehicle_replace(vehicle_db['GWR57xx'], provider, product, blueprint, name, number, rv_num)
        return result
    return False


def lms5xp_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['LMS5XP'])):
        this_vehicle = vehicle_db['LMS5XP'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    nm = re.search('([A-Z])(...)', number.text)
                    if nm:
                        rv_shed_letter = nm.group(1)
                        rv_shed_number = nm.group(2).replace('#', '')
                        rv_loco_number = number.text[-4:]
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), rv_loco_number, '([0-9]{4})(.*)')
                        # Set double chimney for the RSC Jubilee Double Chimney DLC or KWVR Bahamas
                        if 'DoubleChimney' in blueprint.text or '45596 Bahamas' in blueprint.text:
                            rv_num = rv_num[0:11] + 'd' + rv_num[12:]
                        # Set number
                        number.text = rv_num
                        rv_list.append(number.text)
                        rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def bulleidlp_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['BulleidLP'])):
        this_vehicle = vehicle_db['BulleidLP'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    rv_num = this_vehicle[7]
                    if len(rv_num) > 8:
                        number.text = rv_num
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def bulleidrlp_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['BulleidRLP'])):
        this_vehicle = vehicle_db['BulleidRLP'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    nm = re.search('([0-9]{5})', number.text)
                    if nm:
                        rv_num = dcsv_get_num(
                            Path(railworks_path, 'Assets', this_vehicle[3], this_vehicle[4],
                                 this_vehicle[7].replace('\\', '/')), nm.group(1), '([0-9]{5})(.*)')
                        number.text = rv_num
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def dtmaunsell_replace(provider, product, blueprint, name, number, flipped):
    for i in range(0, len(vehicle_db['MaunsellCoach'])):
        this_vehicle = vehicle_db['MaunsellCoach'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text[-4:]
                    if this_vehicle[5] == 'Railvehicles\Passenger\Maunsell Corr\Maunsell-TKD2001-BR(S)-G.xml':
                        number.text = 'S' + number.text[-3:] + 'S'
                    elif this_vehicle[5] == 'Railvehicles\Passenger\Maunsell Corr\Maunsell-CKD2301-BR(S)-G.xml':
                        number.text = 'S' + number.text[-4:] + 'S'
                    elif this_vehicle[5] == 'Railvehicles\Passenger\Maunsell Corr\Maunsell-BCKD2401-BR-G.xml' or \
                            this_vehicle[5] == 'Railvehicles\Passenger\Maunsell Corr\Maunsell-BTKD2101-BR-G.xml':
                        number.text = 'S' + number.text[3:7] + 'S' + number.text[0:3]
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    if flipped.text == '0':
                        flipped.text = '1'
                    else:
                        flipped.text = '0'
                    return True
    return False


def srn15_replace(provider, product, blueprint, name, number):
    for i in range(0, len(vehicle_db['SRN15'])):
        this_vehicle = vehicle_db['SRN15'][i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_orig = number.text
                    number.text = this_vehicle[7]
                    rv_list.append(number.text)
                    rv_pairs.append([rv_orig, number.text])
                    return True
    return False


def vehicle_replace(vehicle_db, provider, product, blueprint, name, number, rv_num):
    for i in range(0, len(vehicle_db)):
        this_vehicle = vehicle_db[i]
        if this_vehicle[0] in provider.text:
            if this_vehicle[1] in product.text:
                bp = re.search(this_vehicle[2], blueprint.text, flags=re.IGNORECASE)
                if bp:
                    provider.text = this_vehicle[3]
                    product.text = this_vehicle[4]
                    blueprint.text = this_vehicle[5]
                    name.text = this_vehicle[6]
                    rv_pairs.append([number.text, rv_num])
                    rv_list.append(rv_num)
                    number.text = rv_num
                    return True
    return False


def user_vehicle_replace(provider, product, blueprint, name):
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
    if values['Replace_Black5'] and black5_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_LMS8F'] and lms8f_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_R04'] and gcr8k_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_GWR57xx'] and gwr57xx_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_LMS5XP'] and lms5xp_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_BulleidLP'] and bulleidlp_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_BulleidRLP'] and bulleidrlp_replace(provider, product, blueprint, name, number):
        return True
    if values['Replace_DTMaunsell'] and dtmaunsell_replace(provider, product, blueprint, name, number, flipped):
        return True
    if values['Replace_SRN15'] and srn15_replace(provider, product, blueprint, name, number):
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
    #if config.get('defaults', 'save_report') == report_opts[1]:
        # User wants a basic report just containing details of the vehicles in the processed scenario
    htmrv = "<h1>Rail vehicle list</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n" \
            "    <tr style=\"text-align: right;\">\n      <th>Consist</th>\n      <th>Provider</th>\n" \
            "      <th>Product</th>\n      <th>Blueprint</th>\n      <th>Name</th>\n      <th>Number</th>\n" \
            "      <th>Loaded</th>\n    </tr>\n  </thead>\n  <tbody>\n"
    #else:
    #    # Full report wanted containing details of the vehicles in the processed scenario as well as original vehicles
    #    htmrv = "<h1>Rail vehicle swap list</h1>\n<table border=\"1\" class=\"dataframe\">\n  <thead>\n" \
    #            "    <tr style=\"text-align: right;\">\n      <th>Consist</th>\n" \
    #            "      <th class=\"input\">Original Provider</th>\n      <th class=\"input\">Original Product</th>\n" \
    #            "      <th class=\"input\">Original Blueprint</th>\n      <th class=\"input\">Original Name</th>\n" \
    #            "      <th class=\"input\">Original Number</th>\n      <th class=\"input\">Loaded</th>\n" \
    #            "      <th>New Provider</th>\n      <th>New Product</th>\n      <th>New Blueprint</th>\n" \
    #            "      <th>New Name</th>\n      <th>New Number</th>\n      <th>Loaded</th>\n    </tr>\n  </thead>\n" \
    #            "  <tbody>\n"
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
    htmrv = htmrv + '  </tbody>\n</table>\n<h3>' + str(len(output_vehicle_list)) + ' vehicles in total in this scenario.</h3>'
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
            sg.Popup('About RSSwapTool-steam',
                     'Tool for swapping rolling stock in Train Simulator (Dovetail Games) scenarios',
                     'Issued under the GNU General Public License - see https://www.gnu.org/licenses/',
                     'Version ' + version_number + ' / ' + version_date,
                     'Copyright 2023 JR McKenzie (jrmknz@yahoo.co.uk)', 'https://github.com/jrmckenzie/RSSwapTool')
        elif event == 'Settings':
            if not config.has_section('defaults'):
                config.add_section('defaults')
            config.set('defaults', 'replace_black5', str(values['Replace_Black5']))
            config.set('defaults', 'replace_lms8f', str(values['Replace_LMS8F']))
            config.set('defaults', 'replace_r04', str(values['Replace_R04']))
            config.set('defaults', 'replace_gwr57xx', str(values['Replace_GWR57xx']))
            config.set('defaults', 'replace_lms5xp', str(values['Replace_LMS5XP']))
            config.set('defaults', 'replace_bulleidlp', str(values['Replace_BulleidLP']))
            config.set('defaults', 'replace_bulleidrlp', str(values['Replace_BulleidRLP']))
            config.set('defaults', 'replace_dtmaunsell', str(values['Replace_DTMaunsell']))
            config.set('defaults', 'replace_srn15', str(values['Replace_SRN15']))
            with open(path_to_config, 'w') as configfile:
                config.write(configfile)
                configfile.close()
            save_report = config.get('defaults', 'save_report', fallback=report_opts[0])
            # The settings button has been pressed, so allow the user to change the RailWorks folder setting
            loclayout = [
                [sg.Text('Settings', justification='c')],
                [sg.Text('Path to RailWorks folder:'),
                 sg.Input(default_text=str(railworks_path), key='RWloc', readonly=True),
                 sg.FolderBrowse(key='RWloc')],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.HSeparator(color='#aaaaaa')],
                [sg.Text("Save a report of vehicles in the scenario"),
                 sg.Combo(report_opts, auto_size_text=True, default_value=save_report, key='save_report', readonly=True,
                          tooltip='You may choose to save a report listing all the rail vehicles (and their numbers)'
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
            config.set('defaults', 'replace_black5', str(values['Replace_Black5']))
            config.set('defaults', 'replace_lms8f', str(values['Replace_LMS8F']))
            config.set('defaults', 'replace_r04', str(values['Replace_R04']))
            config.set('defaults', 'replace_gwr57xx', str(values['Replace_GWR57xx']))
            config.set('defaults', 'replace_lms5xp', str(values['Replace_LMS5XP']))
            config.set('defaults', 'replace_bulleidlp', str(values['Replace_BulleidLP']))
            config.set('defaults', 'replace_bulleidrlp', str(values['Replace_BulleidRLP']))
            config.set('defaults', 'replace_dtmaunsell', str(values['Replace_DTMaunsell']))
            config.set('defaults', 'replace_srn15', str(values['Replace_SRN15']))
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
                                        'RailWorks folder?', 'This application will now exit.')
                        sys.exit()
                    inFile = scenarioPath.parent / Path(str(scenarioPath.stem) + '.xml')
                    inFileW = str(PureWindowsPath(inFile))
                    scenarioPathW = str(PureWindowsPath(scenarioPath))
                    if platform.system() == 'Windows':
                        # Operating system is Microsoft Windows
                        p1 = subprocess.Popen([str(cmd), scenarioPathW, '/xml:' + inFileW], stdout=subprocess.PIPE)
                    elif (platform.system() == 'Linux' and platform.release()[-5:-1] == 'WSL2'):
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
                    elif (platform.system() == 'Linux' and platform.release()[-5:-1] == 'WSL2'):
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
