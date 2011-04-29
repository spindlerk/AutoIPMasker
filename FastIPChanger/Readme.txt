Fast IP Changer - small program written in AutoIT to quickly change your LAN IP address between 4 different configurations.

Any comments, sugestions, bugs, etc should be directed to my sourceforge page
http://sourceforge.net/projects/winipchanger/


Requires Windows XP or Windows Vista
*may also run on other versions of Windows but untested. (please dont email me saying it wont run on your 9x OS)


to use:
open ipchanger.exe
select the adaptor in the "network adaptor" drop down box. if it is not in the list click "Get Adapters" to refresh the list
click on one of the config tabs and make any changes you want.
click the set button on the tab that you would like to enable.
double check the settings are correct by checking the information on the left hand side of the GUI.
click the save button to save the config.

you only need the "Get Adapters" button if you have plugged in a Plug-n-Play device in after running the program,
or if it hasnt detected your NIC.
all it does is re-gets all the network adaptors.

to use the tray options you need to have already setup the configs in the main program
click on the tray icon
navigate to and expand "Enable Config" and select the config of your choice.

Command line options:
ipchanger.exe -p <preset config(1-4)>
(Note uses the last used adapter as specified in the ini)

program has the ability to save the configs for another time and the command line options use the ini file used for saving.


Licence: Free to use and modify to your needs. Just dont put a price on it is all i ask :)




cheers and enjoy