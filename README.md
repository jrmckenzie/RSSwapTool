# RSSwapTool
Tool for swapping rolling stock in Train Simulator (Dovetail Games) scenarios

Many scenarios for Train Simulator use old rolling stock models that are not to the same standard as modern enhancements and new rolling stock from the likes of Armstrong Powerhouse, Fastline Simulations and Vulcan Productions. In some cases, scenarios may use rolling stock that is no longer available to buy. The process of manually updating such scenarios to use newer stock can be extremely time-consuming.

RWSwapTool is a tool for processing scenario files to automate the substitution of older rolling stock for newer versions - mainly for substituting enhancements and new stock from Armstrong Powerhouse (AP) but also some from Fastline Simulations (coal wagons), Just Trains (Voyager) and Vulcan Productions (such as the BR blue Class 47 pack, the Class 86 headcode blinds add-on and other bits and pieces). The main focus is on BR Blue era and later scenarios. It requires a recent version of Python (3.9 or later) to be installed on Windows 10, with the PySimpleGUI add-on installed. It will attempt to automatically re-format the vehicle numbers to be compatible with the relevant AP numbering scheme. A recent version of Train Simulator (railworks.exe) must also be installed.

Note that modifying scenarios in this way (swapping stock) may 'break' the scenario - if swapped vehicles have slightly different physical properties to the originals, they may accelerate and brake differently and have different timings, causing signalling issues as a result. The RSSwapTool program itself may introduce other problems. RSSwapTool is intended to be a tool to help scenario authors update old scenarios. The scenarios output by the tool may well require further manual edits in the Train Simulator Scenario Editor to sort out various issues.

Warning
-------

RSSwapTool is currently under development, in pre-release status and may have various bugs that need to be fixed and features that are incomplete. It may not even work at all. Before using this tool you must accept the terms of the license enclosed within. See also https://www.gnu.org/licenses/gpl-3.0.html.

Installation
============

If you've retrieved the program from github and you've already got Python installed and PySimpleGUI added on, that's it. Python is a free download from https://www.python.org/downloads/release. If you've got Python installed but need to get PySimpleGUI you could try just opening a command prompt and typing "pip install pysimplegui" or consult https://pypi.org/project/PySimpleGUI/ for more information.

If you prefer not to install Python and all it entails, there is a 7zip file you can download for windows 10 with packaged applications RSSwapTool.exe and RSReportTool.exe you just need to double-click to open, which you can get from https://www.dropbox.com/s/zhlir7l9oxfmscw. Just download the whole 7zip file and extract it to your computer. It may not be as up-to-date as the source code here on github, however. 

Usage
=====

Two tools are provided - the main rolling stock swap tool, and a second tool which allows you to examine a scenario and generate a report listing all the rolling stock in the scenario in tabular format. The first time you run either program you will need to specify the location of your RailWorks folder by browsing for it when prompted and pressing "Submit".

#### RSReportTool
To generate the report of the rolling stock within a scenario with **RSReportTool**, without altering the scenario itself, run rs_report.py from the command prompt with "python.exe rs_report.py" or double click it. Then press the "Select scenario file to examine" button and hit "Examine". The "Settings" button will allow you to reconfigure the location of your RailWorks folder if it needs changed. Once a report has been generated, you will be given the option to open it in your web browser. RSReportTool reports will show you the names of the consists, or indicate that they are loose consists if they have no name or driver. The title of the player-driven consist will be in bold and it will say "(Player driven)" after the title. The reports will also show you the asset Provider, Product, Blueprint, Name, Number, and Loaded status. 

You may find some lines in the report are coloured red, this is because the blueprint .bin file for this piece of stock doesn't exist in the expected location. There are two possible reasons for this: you may not have the stock or reskin on your computer; or, you may have the stock but it is packaged with an ".ap" file archive so RSReportTool can't see it. Other tools like TS-Tools (available from Mike Simpson at http://agenetools.com/downloads.html) explain this a bit more and allow you to extract just the .bin files out of the .ap files. If you do that and re-run RSReportTool on the scenario again, it should only show a few red lines for the stock you definitely don't have (or no red lines at all if you have all the stock in the scenario).

#### RSSwapTool
To swap rolling stock with **RSSwapTool**, run main.py from the command prompt with "python.exe main.py" (or just double click main.py if your Windows python is configured to open python scripts automatically). The main screen should then appear. Locate your scenario.bin file by pressing the "Select scenario file to process" button. (It is strongly recommended that you always work on clones of scenarios in case something goes wrong.) Then choose which of the tool's pre-configured swaps you wish to perform on your scenario, by ticking the appropriate box. You can select as many, or as few, as you like. The more you select, the longer it will take. In most cases a scenario should take no more than 30 seconds to process, possibly much less.

The program will save a back-up of the original scenario in a file named after the original file, with the date and time of the substitution operation appended. There is an option (under "Settings") to save a report of all the vehicles in the output scenario without having to run RSReportTool separately. 

You will need to own the rolling stock in question for the scenario to function, this software does not provide any rolling stock it merely edits scenario files for rolling stock you will have to obtain from elsewhere.

The following substitutions are available:

* Replace Mk1 coaches - Tick to enable replacing of Mk1 coaches with AP Mk1 Coach pack vol. 1
* Replace Mk2A-C coaches - Tick to enable replacing of Mk2a coaches with AP Mk2A-C pack
* Replace Mk2D-F coaches - Tick to enable replacing of Mk2e coaches with AP Mk2D-F pack
* Replace FSA/FTA wagons - Tick to enable replacing of FSA and FTA wagons with AP FSA/FTA wagon pack
* Replace HAA wagons - Tick to enable replacing of HAA wagons with mixed wagons from the AP MGR wagon pack
* Replace HHA wagons - Tick to enable replacing of HHA wagons with mixed wagons from the AP HHA wagon pack
* Replace unfitted 21t coal wagons - Tick to enable replacing of unfitted 21t coal wagons with mixed wagons from the Fastline Simulation HTO pack
* Replace fitted 21t coal wagons - Tick to enable replacing of fitted 21t coal wagons with mixed wagons from the Fastline Simulation HTV pack
* Replace VDA wagons - Tick to enable replacing of JL West Highland Line VDA wagons with mixed wagons from the Fastline Simulation VDA pack
* Replace IHH stock - Tick to enable replacing some of the old Iron Horse House (IHH) stock, if you have any (if in doubt leave this unticked)
* Replace User-configured stock - Tick to enable replacing of user-configured stock, contained in file User.csv (leave this unticked unless you have added your own substitutions to User.csv or want to use the supplied examples).
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
* Replace Class 150/2 sets - Tick to enable replacing of Thomson Oovee Class 150s with AP Class 150/2
* Replace Class 156 sets - Tick to enable replacing of Oovee Class 156s with AP Class 156
* Replace Class 158 sets - Tick to enable replacing of North Wales Coast / Settle Carlisle / Fife Circle Class 158s with AP enhanced versions (Cummins, Perkins)
* Replace Class 168, 170, 171 sets - Tick to enable replacing of Thomson Class 170s with AP enhanced versions
* Replace Class 175 sets - Tick to enable replacing of South Wales Coastal / North Wales / Welsh Marches Class 175s with AP enhanced versions (2.0)
* Replace Class 220, 221 sets - Tick to enable replacing of DTG Class 220 Pack XC, Class 221 North Wales Coast and WCML-South sets with JT Advanced Voyagers. Note that, in the case of the DTG WCML-South units, this relies on the Meridian/Voyager Advanced 2019 Overhaul available from https://semaphoresim.com/file/104-meridianvoyager-advanced-2019-overhaul/ and it will replace the WCML-South Avanti livery with a different look - one with Avanti vinyl wraps over the old Virgin livery.   
* Replace Class 319 sets - Tick to enable replacing of BedPanLine 319, DTG Class 319 NSE and RSC Class 319 units with AP versions
* Replace Class 325 sets - Tick to enable replacing of RSC Class 325 pack units with AP enhanced versions
* Replace Class 350 sets - Tick to enable replacing of Portsmouth Direct / Brighton Main Line / WCML Trent Valley / WCML South Class 350 units with AP enhanced versions
* Replace Class 365 sets - Tick to enable replacing of ECMLS London-Peterborough / Class 365 pack Class 365 units with AP enhanced version
* Replace Class 375, 377 sets - Tick to enable replacing of Class 375 Pack / London-Gillingham / Portsmouth Direct / Brighton Main Line / WCML-South Class 375-377 units with AP enhanced versions
* Replace Class 444, 450 sets - Tick to enable replacing of Portsmouth Direct / Class 444 Pack / Brighton Main Line / Guildford Disrict Class 444/450 units with AP enhanced versions
* Replace Class 456 sets - Tick to enable replacing of South London Lines Class 456 units with AP/Waggonz versions
* Replace Class 465 sets - Tick to enable replacing of South London Lines / London-Gillingham / RSC Class 465 Pack Class 465 units with AP enhanced versions

#### Settings

Pressing the 'Settings' button allows you to update the location of RailWorks folder, and change the substitution policy with Class 56, 86, MGR wagons and FSA/FTA wagons. You can choose how to replace a Class 56 where the AP enhancement pack doesn’t include the sector / depot plate of the original. You can choose to replace any Class 86s with headcode blinds with the Vulcan Productions headcode blinds add-on or with plated over headcode box and marker lights. You can choose whether to replace HAA wagons with AP HAA, HBA, HCA, HDA, HFA, HMA & HNA in Blue, Red, or Sector livery. You can choose whether you want Railfreight Distribution or Freightliner FSA/FTA wagons, and whether the containers should be appropriate for 2000s, 2010s, or 2020s era. You can also enable or disable the saving of a report listing all rail vehicles in the scenario and their key details.

#### Notes

Note that in some cases the replacement loco, coach or wagon might come in a variety of different weathering options. The application will choose a random weathering. The HHA, HTO, HTV and VDA wagon replacements will also be slightly randomised within the consist to give a bit of variety - not too many clean new ones though! 

When HAA or VDA wagons are replaced in a driven (not loose) consist, and an HAA or VDA is at the rear of the consist, the application will substitute a variant with a tail lamp facing backwards (regardless of whether the train is driven by the player or AI). 

*(Coding tip: if you wish to modify the behaviour of RSSwapTool while swapping these wagons and randomising the replacements, you can have a look at the top of data_file.py and edit the code. There are some lists of replacement wagons and you can add to or remove from these lists to alter the probability of particular variants appearing.)*

The built-in substitutions are provided by the file Replacements.csv in the "tables" folder, should you wish to add or remove items you may do so at your own risk. There is an option to add your own substitutions ("Replace User-configured stock"). You can do this by configuring the file User.csv in the "tables" folder. A few lines are already entered to give you an idea of the format. (You can open the file in a spreadsheet package or upload and edit it in google sheets). Note that there will be no changes made to rail vehicle numbers for anything included in the User.csv file). If you tick the "Replace User-configured stock" box then any substitutions configured in the User.csv file will be processed. If you've got anything wrong the program may crash, but you can safely delete the contents of the User.csv file and start again if so. 

Note that one of the options is to replace IHH stock. This option is incomplete. None of the old "Iron Horse House" (IHH) stock is available to buy anymore but some scenarios still include it. RSSwapTool will try to replace IHH Class 17, 20, 25, 40, 45 and 47, plus BG and GUV coaches and a 20 ton BR brake van. This will require the DTG BR Blue Pack 01, Vulcan Productions BR Blue reskins for Class 20, 25, 45 and 47. For the time being this option is limited to these IHH items and it may not function very well as the author of RSSwapTool does not own any of the IHH stock to verify. Any Class 17s will be replaced with a Class 20 to provide a loco of similar capabilities.

The "Replace Class 47s" option will replace BR Blue Class 47s with the Vulcan Productions BR Blue Class 47 Pack versions, either with marker lights in a yellow panel or with the domino (whichever was found on the locomotive being replaced). The Class 47s in this pack do not have the orange cant rail or high intensity headlamps which the locomotives being replaced have, so you may only want to select this option if your scenario is from an era where this substitution would be appropriate. Whether the substituted locomotive is of the subclass 47/0, 47/3, 47/4 or 47/7 (and whether it has ETH connections etc on the front) depends on the TOPS number of the locomotive being replaced. Whether or not the locomotive has a grey roof or not also depends on the TOPS number - if the locomotive number in question originally had a grey roof then the swapped-in loco should have a grey roof. Note that only the Brush-manufactured locomotives are provided in the VP pack; this program will give the substituted locomotive the number of the Brush manufactured locomotive closest to the number of the locomotive being replaced, but within the same subclass.

Remember that for various reasons the scenario may need manual editing in Scenario Editor to work properly after the stock is swapped. For example, slightly different physics may change the timing and position of AI at critical points and new signalling issues may manifest themselves. Even loading the scenario into Scenario Editor, making a non-material change (like flipping the orientation of a static wagon in a siding or making a slight change to a vehicle number) may help with some issues as it forces Train Simulator to write out the scenario the way it likes (with its own particular quirks to its xml format) and re-create the ScenarioProperies.xml file.

It is intended that more will be added to this documentation in due course.

Copyright © 2021 JR McKenzie

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>.

The latest version of the source code is available from https://github.com/jrmckenzie/RSSwapTool. 
