RNS-E Modern Map Colours

Reference spreadsheet and Python tool for creating custom colours.map files for the Audi RNS-E navigation system.
Files

    RNSE_Map_Colour_Legend.xlsx Identified RNS-E map colour entries, including Day and Night colour values.

    build_colours_map_gui.py
    Python GUI tool for creating a colours.map file from the spreadsheet.

Requirements

    Python 3
    No additional Python packages required
    PCBBC custom RNS-E firmware and Optional Feature Pack are required to load custom map colours onto the RNS-E.

Usage

    Download both files into the same folder.
    Edit the New Colour values in the spreadsheet.
    Run build_colours_map_gui.py.
    Click Build colours.map.
    Copy the generated colours.map file to your SD card.

Compatibility

Developed and tested on a 193-series RNS-E.

192-series units appear to use different colour assignments and have not been tested.
