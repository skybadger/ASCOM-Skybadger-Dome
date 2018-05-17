Installation instructions for ASCOM CyberDrive Dome Driver:

1) make sure mainstream ASCOM 3.0 package in installed
   in "C:\Program Files\Common Files\ASCOM".  (This is
   the default location).  Edit the bat file otherwise.

2) unpack the zip file into any temporary directory

3) run "InstallCyberDrive.bat" (double click is fine)
  - if first time install, you'll see an error message as
    it tries to unregister any older version.  This is fine.
  - driver and helper dll will be copied
  - driver will be registered with windows
  - driver will be registered with ASCOM, press [OK]
  - process complete, press any key to close window

4) Copy CyberDrive launch shortcut wherever...

5) Delete temporary directory from step 2

CyberDrive can be launched manually, or found in the ASCOM chooser.
