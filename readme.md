# SDPlus CSC Call Classifier
Made for North Bristol Trust Back Office Team, whilst working as a Senior Clinical Systems Analyst, to help analyse logged issue Service Level Agreement resolution times.

## Background:
This call goes through the Back Office Third Party/CSC in Manage Engine's Service Desk Plus helpdesk software, through every item and conversation in each item, will then output results on screen (i.e. which party responded last), copy to the clipboard, and will open an Excel sheet (in the windows temporary folder) of the results.

## Prerequisite
Install "Visual C++ Redistributable for Visual Studio 2015 x86.exe" (on 32-bit, or x64 on 64-bit) which allows Python 3.5 dlls to work, found here:
https://www.microsoft.com/en-gb/download/details.aspx?id=48145

## Installation and Running
Just double click `sdplus_classify_calls.exe` from any location to run the program.

### Notes
This program communicates with the Service Deskplus API via an sdplus_api_technician_key which can be obtained via the sdplus section:
Admin, Assignees, Edit Assignee (other than yourself), Generate API Key.
It will look for an API key in a windows variable under name "SDPLUS_ADMIN". You can set this on windows with:
`setx SDPLUS_ADMIN <insert your own SDPLUS key here>`
in a command line.
If not found, the program will exit.

_Written by:_  
_Simon Crouch, late 2016 in Python 3.5_