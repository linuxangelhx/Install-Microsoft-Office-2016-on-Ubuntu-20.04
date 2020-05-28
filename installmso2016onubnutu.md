Requirements
-

To install Microsoft Office you have to install PlayOnLinux, samba, winbind and winetricks for it to properly work.

The first step is to update & upgrade all packages.

```
$sudo apt update && sudo apt upgrade
```

then

```
$sudo apt-get install playonlinux samba winbind winetricks unrar
```

Download the MSO 2016 installer files (either DVD/folder files), in the 32 bits version. Even if you're under Ubuntu 64, we'll use a 32 bits wine installation.


**You can download Microsfot Office 2016 32bits from this link** [MSO 2016 32 Bits Installer]( https://doc1034docs.googleusercontent.com/docs/securesc/k0o5m51bvqtq8pl2aqf2e0un3aoamjpa/nqcs4qpgrnfvbim8ksl7q67ojug4ci9b/1590678450000/03322087261072322119/05396892547377328719/13bxXRGJTLJCBB_OOejXTT5TTjETZ59VD?e=download&authuser=0&nonce=4g2bv5om463tk&user=05396892547377328719&hash=dpv5rc21eehh8r3u3vrr9fout45nsemc)

* *I also installed this version and I'm having some issues with Outlook but I don't use it. Everything else works fine.* *

When you have the MSO files downloaded, use `unrar x` command in the console to decompress the file.

```
$unrar x filename.rar
```

**Note: I recommend to rename the file to a short name like office16.rar or whatever you want, becasuse when you decompress, it will be more easy.**


Now the MSO Installer is downloaded, you should save the next script as mso2016.sh


* *The Script* *

```
#!/bin/bash

# CHANGELOG
# Version 0.0.1 by N0rbert - with Wine 3.0 it crashes on Welcome screen when trying
# to call "unimplemented function KERNEL32.dll.FindNLSStringEx called in 32-bit code"
# So we need Wine 3.4+ here, but I can't run this script with it on my Ubuntu 16.04 LTS.
#
# based on Version 1.1.0.0 [Quentin Pâris and Eduardo Lucio (Questor)] (2018-02-17 15-24) - Paris/Brazil
# for office2007pro (playonlinux://www.playonlinux.com/repository/download.php?id=2665)

# >>>>>>>>>>
# Under BSD License!

# Copyright (c) 2018, Quentin Pâris, Eduardo Lucio and N0rbert
# All rights reserved.

# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#     * Redistributions of source code must retain the above copyright
#       notice, this list of conditions and the following disclaimer.
#     * Redistributions in binary form must reproduce the above copyright
#       notice, this list of conditions and the following disclaimer in the
#       documentation and/or other materials provided with the distribution.
#     * Neither the name of the free software community nor the
#       names of its contributors may be used to endorse or promote products
#       derived from this software without specific prior written permission.

# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL Quentin Pâris and Eduardo Lucio BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
# SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

# <<<<<<<<<<

# NOTE: Complete liste of references! By Questor
# https://github.com/PlayOnLinux/POL-POM-4/blob/master/lib/wine.lib
# https://www.playonlinux.com/en/app-2665-Microsoft_Office_2016.html
# http://wiki.playonlinux.com/index.php/Scripting_-_Chapter_11:_List_of_Functions
# http://wiki.playonlinux.com/index.php/Components_and_Functions
# https://www.playonlinux.com/repository/source.php?script=822
# https://github.com/PlayOnLinux/POL-POM-4/blob/367e50865168b5b931611fa33b0c1d8426143a90/lib/scripts.lib
# https://github.com/PlayOnLinux/POL-POM-4/blob/367e50865168b5b931611fa33b0c1d8426143a90/lib/wine.lib
# https://askubuntu.com/questions/879304/wine-2-0-says-it-supports-office-2016-how-do-i-actually-install-it
# https://github.com/RobertJames/playonlinux/blob/75ef37523f299255a539a89b63dc87d7afc096d4/template.POL

# N0rbert's links:
# https://bugs.winehq.org/show_bug.cgi?id=41911 (really we need Wine 3.4+ to fix FindNLSStringEx errors)

# >>>>>>>>>>
# Initialization!

[ "$PLAYONLINUX" = "" ] && exit 0
source "$PLAYONLINUX/lib/sources"

PREFIX="office2016"

WINEVERSION="3.4"

TITLE="Microsoft Office 2016"

# NOTE: Images to use! By Questor
POL_GetSetupImages "https://i.imgur.com/licFVuF.png" "https://i.imgur.com/ff6PkEZ.png" "$TITLE"

POL_SetupWindow_Init
POL_SetupWindow_SetID 3064

POL_SetupWindow_presentation "$TITLE" "Microsoft" "http://www.microsoft.com" "Quentin Pâris, Eduardo Lucio and N0rbert" "$PREFIX"

POL_Debug_Init

# <<<<<<<<<<

# >>>>>>>>>>
# Perform some validations!

POL_RequiredVersion 4.2.10 || POL_Debug_Fatal "$TITLE won't work with $APPLICATION_TITLE $VERSION!\nPlease update!"

if [ "$POL_OS" = "Linux" ]; then
    wbinfo -V || POL_Debug_Fatal "Please install winbind before installing $TITLE!"
fi

if [ "$POL_OS" = "Mac" ]; then

    # NOTE: Samba support! By Quentin Pâris
    POL_Call POL_GetTool_samba3
    source "$POL_USER_ROOT/tools/samba3/init"

fi

POL_Wine_WaitBefore "$TITLE"
[ "$CDROM" ] && cd "$CDROM"

if [ ! "$(file $SetupIs | grep 'x86-64')" = "" ]; then
    POL_Debug_Fatal "$(eval_gettext "The 64bits version is not compatible! Sorry!")";
fi

# NOTE: Check if "winetricks" is present! By Questor
winetricks -V || POL_Debug_Fatal "Please install winetricks before installing $TITLE!"

POL_System_SetArch "x86"
POL_SetupWindow_InstallMethod "LOCAL,DVD"

if [ "$INSTALL_METHOD" = "DVD" ]; then
    POL_SetupWindow_cdrom
    POL_SetupWindow_check_cdrom "x86/setup.exe" "setup.exe"
    SetupIs="$CDROM_SETUP"
    cd "$CDROM"
else
    POL_SetupWindow_browse "$(eval_gettext 'Please select the setup file to run!')" "$TITLE"
    SetupIs="$APP_ANSWER"
fi

# <<<<<<<<<<

# >>>>>>>>>>
# Prepare resources for installation!

# NOTE: Install wine version if isn't available. This is necessary because
# even though "POL_Wine_PrefixCreate" solves this, we end up having
# problems when the required version is not available and it tries to
# install it! Questor
# [Ref.: https://github.com/PlayOnLinux/POL-POM-4/blob/master/lib/wine.lib]
POL_Wine_InstallVersion "$WINEVERSION"

POL_Wine_SelectPrefix "$PREFIX"
POL_Wine_PrefixCreate "$WINEVERSION"

Set_OS "win7"

# Fix black windows (added by N0rbert)
POL_Wine_Direct3D "MaxVersionGL" "30002"

# <<<<<<<<<<

# >>>>>>>>>>
# Install!

# NOTE: Installs office! By Questor
POL_Wine "$SetupIs"
POL_Wine_WaitExit "$TITLE"

# <<<<<<<<<<

# >>>>>>>>>>
# Prepare resources for applications!

# NOTE: Uses native special version of "riched20" installed by Office 2016!
# Fix the "black fields" issue! By Questor
# [Ref.: http://forum.winehq.org/viewtopic.php?f=8&t=23126&p=95555#p95555]
POL_Wine_OverrideDLL "native,builtin" "riched20"

# NOTE: Fix "move and change the window size (maximize/minimize/restore/resize
# etc...) bugs"! By Questor
# [Ref.: https://bugs.winehq.org/show_bug.cgi?id=44552]
Set_Managed "Off"

# <<<<<<<<<<

# >>>>>>>>>>
# Create shortcuts, entries to extensions and finalize!

# NOTE: Create shortcuts! By Questor
POL_Shortcut "WINWORD.EXE" "Microsoft Word 2016" "" "" "Office;WordProcessor;"
POL_Shortcut "EXCEL.EXE" "Microsoft Excel 2016" "" "" "Office;Spreadsheet;"
POL_Shortcut "POWERPNT.EXE" "Microsoft Powerpoint 2016" "" "" "Office;Presentation;"

# NOTE: No category for collaborative work? By Quentin Pâris
POL_Shortcut "ONENOTE.EXE" "Microsoft OneNote 2016" "" "" "Network;InstantMessaging;"

# NOTE: "Calendar;ContactManagement;"? By Quentin Pâris
POL_Shortcut "OUTLOOK.EXE" "Microsoft Outlook 2016" "" "" "Network;Email;"

# NOTE: Add an entry to PlayOnLinux's extension file. If the entry already
# exists, it will replace it! By Questor
# [Ref.: https://github.com/PlayOnLinux/POL-POM-4/blob/master/lib/playonlinux.lib]
POL_Extension_Write doc "Microsoft Word 2016"
POL_Extension_Write docx "Microsoft Word 2016"
POL_Extension_Write xls "Microsoft Excel 2016"
POL_Extension_Write xlsx "Microsoft Excel 2016"
POL_Extension_Write ppt "Microsoft Powerpoint 2016"
POL_Extension_Write pptx "Microsoft Powerpoint 2016"

if [ "$POL_OS" = "Mac" ]; then
    POL_Shortcut_InsertBeforeWine "Microsoft Word 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
    POL_Shortcut_InsertBeforeWine "Microsoft Excel 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
    POL_Shortcut_InsertBeforeWine "Microsoft Powerpoint 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
    POL_Shortcut_InsertBeforeWine "Microsoft OneNote 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
    POL_Shortcut_InsertBeforeWine "Microsoft Outlook 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
fi

POL_SetupWindow_message "$(eval_gettext '$TITLE has been installed successfully!\n\nThanks!\nBy Quentin Pâris, Eduardo Lucio and N0rbert')" "$TITLE"
POL_SetupWindow_Close

# <<<<<<<<<<

exit 0
One should save it as pol-mso2016.sh and then load it into PlayOnLinux by selecting Tools -> Run a local script and then follow instructions on the screen.

share  improve this answer  follow 
edited Aug 8 '18 at 8:35
answered Jun 8 '18 at 18:15

N0rbert
43.9k1212 gold badges100100 silver badges215215 bronze badges
Doesn't the official Wine PPA support Wine 3.4 on Ubuntu 16.04? If so, you could use that Wine version, and set the script to use the 'System' Wine Version instead of defining a certain version number. Also you might want to make sure PlayOnLinux is the latest version via the PPA, in case it's the version of POL preventing the usage of Wine 3.4 in Ubuntu 16.04. – Dominic Hayes Sep 8 '18 at 10:17
Could you host it in a gist or somesuch? It's a little hard to copy from the webpage here. – retorquere Sep 10 '18 at 19:31
1
Testing in 18.10 with POL 4.3.2 and wine 3.4, the setup.exe installation failed – user47206 Nov 20 '18 at 11:41 
add a comment

3

What worked for me was translating intructions from WineHQ (here) and here to use with PlayonLinux, with Wine 3.4.

Create a new virtual drive
Install msxml6 and riched20 from tab "Install Components"
Configure wine to use msxml6 and riched20 (native, builtin)
Copy the AppvIsvSubsystems32.dll and the C2R32.dll from: /Program Files/Common Files/Microsoft Shared/ClickToRun/ to: /Program Files/Microsoft Office/root/Office16/
Edit registry by adding a DWORD named HKCU/Software/Wine/Direct2D/max_version_factory, set to zero.
share  improve this answer  follow 
edited Sep 24 '18 at 9:19
answered Sep 24 '18 at 9:07

Kevin Zarca
13144 bronze badges
1
It didn't work for me in 18.10. Step 4 is to be done after the installation of msoffice I imagine, otherwise that file to be moved is absent. But executing msoffice setup.exe fails. – user47206 Nov 20 '18 at 11:37
I had to copy the file "AppvIsvStream32.dll" as well in step 4. Also, as pointed out by user47206, running the installer has to be done before step 4. – Mohamed Y. Mar 16 at 10:48
add a comment

2

For me the script from above (originally by Quentin Pâris, Eduardo Lucio and N0rbert) worked to install Office 2016 Pro Plus under Ubuntu 18.04, with the only changes:

wine version 3.0
msxml6 in addition to riched20 as overrideDLL
So far Word, Excel, and Powerpoint work, while OneNote and Outlook crashed (don't use them so I didn't even try to fix). Activation worked also, for this open Word -> File -> Account and enter the license key. Here the script, proceed as described above (save as .sh and run under PlayOnLinux -> Tools -> Run local script):

#!/bin/bash

# CHANGELOG
# Version 0.0.1 by N0rbert - with Wine 3.0 it crashes on Welcome screen when trying
# to call "unimplemented function KERNEL32.dll.FindNLSStringEx called in 32-bit code"
# So we need Wine 3.4+ here, but I can't run this script with it on my Ubuntu 16.04 LTS.
#
# based on Version 1.1.0.0 [Quentin Pâris and Eduardo Lucio (Questor)] (2018-02-17 15-24) - Paris/Brazil
# for office2007pro (playonlinux://www.playonlinux.com/repository/download.php?id=2665)

# >>>>>>>>>>
# Under BSD License!

# Copyright (c) 2018, Quentin Pâris, Eduardo Lucio and N0rbert
# All rights reserved.

# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#     * Redistributions of source code must retain the above copyright
#       notice, this list of conditions and the following disclaimer.
#     * Redistributions in binary form must reproduce the above copyright
#       notice, this list of conditions and the following disclaimer in the
#       documentation and/or other materials provided with the distribution.
#     * Neither the name of the free software community nor the
#       names of its contributors may be used to endorse or promote products
#       derived from this software without specific prior written permission.

# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL Quentin Pâris and Eduardo Lucio BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
# SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

# <<<<<<<<<<

# NOTE: Complete liste of references! By Questor
# https://github.com/PlayOnLinux/POL-POM-4/blob/master/lib/wine.lib
# https://www.playonlinux.com/en/app-2665-Microsoft_Office_2016.html
# http://wiki.playonlinux.com/index.php/Scripting_-_Chapter_11:_List_of_Functions
# http://wiki.playonlinux.com/index.php/Components_and_Functions
# https://www.playonlinux.com/repository/source.php?script=822
# https://github.com/PlayOnLinux/POL-POM-4/blob/367e50865168b5b931611fa33b0c1d8426143a90/lib/scripts.lib
# https://github.com/PlayOnLinux/POL-POM-4/blob/367e50865168b5b931611fa33b0c1d8426143a90/lib/wine.lib
# https://askubuntu.com/questions/879304/wine-2-0-says-it-supports-office-2016-how-do-i-actually-install-it
# https://github.com/RobertJames/playonlinux/blob/75ef37523f299255a539a89b63dc87d7afc096d4/template.POL

# N0rbert's links:
# https://bugs.winehq.org/show_bug.cgi?id=41911 (really we need Wine 3.4+ to fix FindNLSStringEx errors)

# >>>>>>>>>>
# Initialization!

[ "$PLAYONLINUX" = "" ] && exit 0
source "$PLAYONLINUX/lib/sources"

PREFIX="office2016"

WINEVERSION="3.0"

TITLE="Microsoft Office 2016"

# NOTE: Images to use! By Questor
POL_GetSetupImages "https://i.imgur.com/licFVuF.png" "https://i.imgur.com/ff6PkEZ.png" "$TITLE"

POL_SetupWindow_Init
POL_SetupWindow_SetID 3064

POL_SetupWindow_presentation "$TITLE" "Microsoft" "http://www.microsoft.com" "Quentin Pâris, Eduardo Lucio and N0rbert" "$PREFIX"

POL_Debug_Init

# <<<<<<<<<<

# >>>>>>>>>>
# Perform some validations!

POL_RequiredVersion 4.2.10 || POL_Debug_Fatal "$TITLE won't work with $APPLICATION_TITLE $VERSION!\nPlease update!"

if [ "$POL_OS" = "Linux" ]; then
    wbinfo -V || POL_Debug_Fatal "Please install winbind before installing $TITLE!"
fi

if [ "$POL_OS" = "Mac" ]; then

    # NOTE: Samba support! By Quentin Pâris
    POL_Call POL_GetTool_samba3
    source "$POL_USER_ROOT/tools/samba3/init"

fi

POL_Wine_WaitBefore "$TITLE"
[ "$CDROM" ] && cd "$CDROM"

if [ ! "$(file $SetupIs | grep 'x86-64')" = "" ]; then
    POL_Debug_Fatal "$(eval_gettext "The 64bits version is not compatible! Sorry!")";
fi

# NOTE: Check if "winetricks" is present! By Questor
winetricks -V || POL_Debug_Fatal "Please install winetricks before installing $TITLE!"

POL_System_SetArch "x86"
POL_SetupWindow_InstallMethod "LOCAL,DVD"

if [ "$INSTALL_METHOD" = "DVD" ]; then
    POL_SetupWindow_cdrom
    POL_SetupWindow_check_cdrom "x86/setup.exe" "setup.exe"
    SetupIs="$CDROM_SETUP"
    cd "$CDROM"
else
    POL_SetupWindow_browse "$(eval_gettext 'Please select the setup file to run!')" "$TITLE"
    SetupIs="$APP_ANSWER"
fi

# <<<<<<<<<<

# >>>>>>>>>>
# Prepare resources for installation!

# NOTE: Install wine version if isn't available. This is necessary because
# even though "POL_Wine_PrefixCreate" solves this, we end up having
# problems when the required version is not available and it tries to
# install it! Questor
# [Ref.: https://github.com/PlayOnLinux/POL-POM-4/blob/master/lib/wine.lib]
POL_Wine_InstallVersion "$WINEVERSION"

POL_Wine_SelectPrefix "$PREFIX"
POL_Wine_PrefixCreate "$WINEVERSION"

Set_OS "win7"

# Fix black windows (added by N0rbert)
POL_Wine_Direct3D "MaxVersionGL" "30002"

# <<<<<<<<<<

# >>>>>>>>>>
# Install!

# NOTE: Installs office! By Questor
POL_Wine "$SetupIs"
POL_Wine_WaitExit "$TITLE"

# <<<<<<<<<<

# >>>>>>>>>>
# Prepare resources for applications!

# NOTE: Uses native special version of "riched20" installed by Office 2016!
# Fix the "black fields" issue! By Questor
# [Ref.: http://forum.winehq.org/viewtopic.php?f=8&t=23126&p=95555#p95555]
POL_Wine_OverrideDLL "native,builtin" "riched20"
POL_Wine_OverrideDLL "native,builtin" "msxml6"

# NOTE: Fix "move and change the window size (maximize/minimize/restore/resize
# etc...) bugs"! By Questor
# [Ref.: https://bugs.winehq.org/show_bug.cgi?id=44552]
Set_Managed "Off"

# <<<<<<<<<<

# >>>>>>>>>>
# Create shortcuts, entries to extensions and finalize!

# NOTE: Create shortcuts! By Questor
POL_Shortcut "WINWORD.EXE" "Microsoft Word 2016" "" "" "Office;WordProcessor;"
POL_Shortcut "EXCEL.EXE" "Microsoft Excel 2016" "" "" "Office;Spreadsheet;"
POL_Shortcut "POWERPNT.EXE" "Microsoft Powerpoint 2016" "" "" "Office;Presentation;"

# NOTE: No category for collaborative work? By Quentin Pâris
POL_Shortcut "ONENOTE.EXE" "Microsoft OneNote 2016" "" "" "Network;InstantMessaging;"

# NOTE: "Calendar;ContactManagement;"? By Quentin Pâris
POL_Shortcut "OUTLOOK.EXE" "Microsoft Outlook 2016" "" "" "Network;Email;"

# NOTE: Add an entry to PlayOnLinux's extension file. If the entry already
# exists, it will replace it! By Questor
# [Ref.: https://github.com/PlayOnLinux/POL-POM-4/blob/master/lib/playonlinux.lib]
POL_Extension_Write doc "Microsoft Word 2016"
POL_Extension_Write docx "Microsoft Word 2016"
POL_Extension_Write xls "Microsoft Excel 2016"
POL_Extension_Write xlsx "Microsoft Excel 2016"
POL_Extension_Write ppt "Microsoft Powerpoint 2016"
POL_Extension_Write pptx "Microsoft Powerpoint 2016"

if [ "$POL_OS" = "Mac" ]; then
    POL_Shortcut_InsertBeforeWine "Microsoft Word 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
    POL_Shortcut_InsertBeforeWine "Microsoft Excel 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
    POL_Shortcut_InsertBeforeWine "Microsoft Powerpoint 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
    POL_Shortcut_InsertBeforeWine "Microsoft OneNote 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
    POL_Shortcut_InsertBeforeWine "Microsoft Outlook 2016" "source \"$POL_USER_ROOT/tools/samba3/init\""
fi

POL_SetupWindow_message "$(eval_gettext '$TITLE has been installed successfully!\n\nThanks!\nBy Quentin Pâris, Eduardo Lucio and N0rbert')" "$TITLE"
POL_SetupWindow_Close

# <<<<<<<<<<

exit 0

```

Once you have it ready, open PlayOnLinux and then load the script by selecting Tools -> Run a local script and then follow instructions on the screen.

**_PlayOnLinux Wizard Warning appear, select next, check the** `I Agree` **box and then click on next._**

Automatically the MSO installation wizard will be displayed, select next and then select `Use a setup file in my computer` and click on next again.

In the next window, click on **browse**, find and select **setup.exe** from decompressed installation folder were it is located.

A successful message: click on Next and the install window will be closed.

Now you can close PlayOnLinux.

***Note: To launch any office suite shorcuts on your desktop, just right click and choose*** `Allow launching`, ***now you can run the MSO apps.***

:)


Sources:

```
hoabandoblog@gmail.com (Google Drive Folder)
https://askubuntu.com/questions/975104/how-do-i-install-ms-office-2016-on-playonlinux (Script)
``` 
