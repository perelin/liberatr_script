LiberatR
========

## Liberate your iTunes ratings!
Python script for Windows to backup and restore iTunes ratings.

## Approach
The ratings (and playcounts) are copied directly into to the MP3 files, 
embeded as POPM frames according to ID3 v2.4.0 specification http://id3.org/id3v2.4.0-frames
The backup will be read from the written frame.

## Dependencies
The script makes use of the iTunes COM API, Pythons win32com extension and the mutagen Python multimedia tagging library.
They all need to be available in order to run the script.

## Usage
When all dependencies are met start the script with
python liberatr -h
to get an overview of the available options
