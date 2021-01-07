# RSSwapTool
Tool for swapping rolling stock in Train Simulator (Dovetail Games) scenarios

Many scenarios for Train Simulator use old rolling stock models that are not to the same standard as modern enhancements and new rolling stock from the likes of Armstrong Powerhouse, Fastline Simulations and Vulcan Productions. In some cases, scenarios may use rolling stock that is no longer available to buy. The process of manually updating such scenarios to use newer stock can be extremely time-consuming.

RWSwapTool is a tool for processing scenario files to automate the substitution of older rolling stock for newer versions - mainly for substituting enhancements and new stock from Armstrong Powerhouse (AP) but also some from Fastline Simulations (coal wagons) and Vulcan Productions (such as the BR blue Class 47 pack, the Class 86 headcode blinds add-on and other bits and pieces). It requires a recent version of Python (3.9 or later) to be installed on Windows 10, with the PySimpleGUI add-on installed. It will attempt to automatically re-format the vehicle numbers to be compatible with the relevant AP numbering scheme. A recent version of Train Simulator (railworks.exe) must also be installed.

Note that modifying scenarios in this way (swapping stock) may 'break' the scenario - if swapped vehicles have slightly different physical properties to the originals, they may accelerate and brake differently and have different timings, causing signalling issues as a result. The RSSwapTool program itself may introduce other problems. RSSwapTool is intended to be a tool to help scenario authors update old scenarios. The scenarios output by the tool may well require further manual edits in the Train Simulator Scenario Editor to sort out various issues.

Warning
-------

RSSwapTool is currently under development, in pre-release status and may have various bugs that need to be fixed and features that are incomplete. It may not even work at all. Before using this tool you must accept the terms of the license enclosed within. See also https://www.gnu.org/licenses/gpl-3.0.html.

Installation
============

If you've retrieved the program from github and you've already got Python installed and PySimpleGUI added on, that's it. Python is a free download from https://www.python.org/downloads/release. If you've got Python installed but need to get PySimpleGUI you could try just opening a command prompt and typing "pip install pysimplegui" or consult https://pypi.org/project/PySimpleGUI/ for more information.

Usage
=====

Run main.py from the command prompt with "python.exe main.py" (or just double click main.py if your Windows python is configured to open python scripts automatically).

The first time you run the program you will need to specify the location of your RailWorks folder by browsing for it and pressing "Submit". It is strongly recommeneded that you work on clones of scenarios in case something goes wrong. If you clone the scenario via the Train Simulator (TS) "Build" screens, it will tell you the exact path to your scenario file.

Locate your scenario.bin file by pressing the "Select scenario file to process" button. Then choose which of the tool's pre-configured swaps you wish to perform on your scenario, by ticking the appropriate box. You can select as many, or as few, as you like. The more you select, the longer it will take. In most cases a scenario should take no more than 30 seconds to process, possibly much less.

The program will save a back-up of the original scenario in a file named after the original file, with the date and time of the substitution operation appended. There is an option (under "Settings") to save a report of all the vehicles in the output scenario. 

You will need to own the rolling stock in question for the scenario to function, this software does not provide any rolling stock it merely edits scenario files for rolling stock you will have to obtain from elsewhere.

The following substitutions are available:

* Replace Mk1 coaches - Tick to enable replacing of Mk1 coaches with AP Mk1 Coach pack vol. 1
* Replace Mk2A-C coaches - Tick to enable replacing of Mk2a coaches with AP Mk2A-C pack
* Replace Mk2D-F coaches - Tick to enable replacing of Mk2e coaches with AP Mk2D-F pack
* Replace FSA/FTA wagons - Tick to enable replacing of FSA and FTA wagons with AP FSA/FTA wagon pack
* Replace HAA wagons - Tick to enable replacing of HAA wagons with mixed wagons from the AP MGR wagon pack
* Replace unfitted 21t coal wagons - Tick to enable replacing of unfitted 21t coal wagons with mixed wagons from the Fastline Simulation HTO pack
* Replace fitted 21t coal wagons - Tick to enable replacing of fitted 21t coal wagons with mixed wagons from the Fastline Simulation HTV pack
* Replace VDA wagons - Tick to enable replacing of JL West Highland Line VDA wagons with mixed wagons from the Fastline Simulation VDA pack
* Replace IHH stock - Tick to enable replacing of old Iron Horse House (IHH) stock, if you have any (if in doubt leave this unticked)
* Replace User-configured stock - Tick to enable replacing of user-configured stock, contained in file User.csv (leave this unticked unless you have added your own substitutions to User.csv)
* Replace Class 31s - Replace Class 31s with AP enhancement pack equivalent
* Replace Class 37s - Replace Class 37s with AP equivalent
* Replace Class 40s - Replace DT Class 40s with AP/RailRight equivalent
* Replace Class 47s - Replace BR Blue Class 47s with Vulcan Productions BR Blue Class 47 Pack versions
* Replace Class 50s - Replace MeshTools Class 50s with AP equivalent
* Replace Class 56s - Replace RSC Class 56 Railfreight Sectors with AP enhancement pack equivalent
* Replace Class 66s - Replace Class 66s with AP enhancement pack equivalent
* Replace Class 67s - Replace Class 67s with AP enhancement pack equivalent
* Replace Class 68s - Replace Class 68s with AP enhancement pack equivalent
* Replace Class 86s - Replace Class 86s with AP enhancement pack equivalent
* Replace HST sets - Tick to enable replacing of HST sets with AP enhanced versions (Valenta, MTU, VP185)
* Replace Class 91 EC sets - Tick to enable replacing of Class 91 East Coast sets with AP enhanced versions
* Replace Class 101 sets - Tick to enable replacing of retired RSC Class101Pack with RSC BritishRailClass101 sets - note there is no BR White livery available, these will be replaced with BR Blue
* Replace Class 158 sets - Tick to enable replacing of North Wales Coast / Settle Carlisle / Fife Circle Class 158s with AP enhanced versions (Cummins, Perkins)

There is an option to add your own substitutions ("Replace User-configured stock"). You can do this be configuring the file User.csv in the "tables" folder. A few lines are already entered to give you an idea of the format. (You can open the file in a spreadsheet package or upload and edit it in google sheets). Note that there will be no changes made to rail vehicle numbers for anything included in the User.csv file). If you tick the "Replace User-configured stock" box then any substitutions configured in the User.csv file will be processed. If you've got anything wrong the program may crash, but you can safely delete the User.csv file and start again if so. The built-in substitutions are provided by the file Replacements.csv in the "tables" folder, should you wish to add or remove items you may do so. 

Note that one of the options is to replace IHH stock. This option is incomplete. None of the old "Iron Horse House" (IHH) stock is available to buy anymore but some scenarios still include it. RSSwapTool will try to replace IHH Class 20, 25, 40, 45 and 47, plus BG and GUV coaches and a 20 ton BR brake van. This will require the DTG BR Blue Pack 01, Vulcan Productions BR Blue reskins for Class 20, 25, 45 and 47. For the time being this option is limited to these IHH items and it may not function very well as the author of RSSwapTool does not own any of the IHH stock to verify.

The "Replace Class 47s" option will replace BR Blue Class 47s with the Vulcan Productions BR Blue Class 47 Pack versions, either with marker lights in a yellow panel or with the domino (whichever was found on the locomotive being replaced). The Class 47s in this pack do not have the orange cant rail or high intensity headlamps which the locomotives being replaced have, so you may only want to select this option if your scenario is from an era where this substitution would be appropriate. Whether the substituted locomotive is of the subclass 47/0, 47/3, 47/4 or 47/7 (and whether it has ETH connections etc on the front) depends on the TOPS number of the locomotive being replaced. Whether or not the locomotive has a grey roof or not also depends on the TOPS number - if the locomotive number in question originally had a grey roof then the swapped-in loco should have a grey roof. Note that only the Brush-manufactured locomotives are provided in the VP pack; this program will give the substituted locomotive the number of the Brush manufactured locomotive closest to the number of the locomotive being replaced.

There is also a "Settings" button you can press to update some settings (e.g. location of RailWorks folder, substitution behaviour with Class 56 and 86, saving of a report listing all rail vehicles in the scenario and their key details).
