# WSUS
Script for Managing WSUS

There are a few different scripts included in this repository.
1. Fix-WSUS just does the typical steps of resetting the PCs connection to WSUS. USeful when troubleshooting why a PC is not checking in.
2. WSUS-GUI is a work in progress version of the WSUS control script that will allow for changes to be made in an interface
3. WSUS is the main script and just approves updates older than X months, while denying specific updates and importing a .csv file and denying all updates in that list as well.
