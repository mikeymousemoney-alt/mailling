import os  # Miscellaneous operating system interfaces
import logging  # logging
import time
import datetime as dt  # Basic date and time types
#import win32com.client  # to access outlook application
import json # read data from json files
from pathlib import Path

def main(is_frozen): 
    try:
        if (is_frozen):
            filename = './config.txt'
            config = 3
        else:
            filename = "Vector_Issue\\..\\..\\config.txt"
            config = 1
        open(filename)
    except:    
        try:
            filename = "src\\..\\config.txt"
            config = 2
            open(filename)
        except:
            print("config.txt not found")
            logging.error("config.txt not found")
            exit()

    dict1 = {}

    with open(filename) as fh:
    
        # get rid of empty lines
        lines = (line.rstrip() for line in fh) 
        lines = list(line for line in lines if line)
    
        for line in lines:
            # remove comments
            #print("line before split: " + line)
            line = line.split('////', 1)[0]
            if line.rstrip():
                #print("line without comment: " + line)
                # read key and value through a split at :
                command, description = line.strip().split(':', 1)
    
                dict1[command] = description.strip()
            #print()

    try:
        if config == 1 :
            out_filename = "Vector_Issue/config.json"
            out_file = open(out_filename, "w") 
        if config == 2 :
            out_filename = "src/Vector_Issue/config.json"
            out_file = open(out_filename, "w")
        else:
            out_filename = "temp/config.json"
            #for executable the folder could not exist
            folder = Path(out_filename)
            folder.parent.mkdir(parents=True, exist_ok=True)
            out_file = open(out_filename, "w")

    except:
        print("make_config.py - open file not possible:", out_filename)
        return ''

     # Konvertiere MIA_FILTER_NOSECURITY in eine Liste
    if 'MIA_FILTER_NOSECURITY' in dict1:
        dict1['MIA_FILTER_NOSECURITY'] = [pattern.strip() for pattern in dict1['MIA_FILTER_NOSECURITY'].split(',')]

    json.dump(dict1, out_file, indent = 4, sort_keys = False)
    out_file.close()
    return out_filename