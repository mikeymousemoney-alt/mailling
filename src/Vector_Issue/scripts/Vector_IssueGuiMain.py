#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
================================
Vector_IssueGui
================================

The Vector_IssueGui program will execute a gui. 


'''
__author__ = 'Dominik Schubert'
__copyright__ = 'Copyright 2023, Marquardt'
__credits__ = []
__license__ = 'Marquardt'
__version__ = '0.0.0'
__maintainer__ = 'Dominik Schubert'
__email__ = 'dominik.schubert@marquardt.com'
__status__ = 'Development'

import Vector_Issue.Vector_Issue as Vector_Issue
import subprocess
import time
import argparse
import logging


def my_func_that_returns_a_parser():
       
    # Load cammand line options
    parser = argparse.ArgumentParser(
                    prog = 'Vector_IssueGui',
                    description = __doc__ ,
                    epilog = __copyright__ , 
                    formatter_class = argparse.RawTextHelpFormatter )

    parser.add_argument(    nargs='*'
                            , type    = str
                            , default = []                
                            , dest    = "plugins"          
                            , help    = "make lots of noise [default]"    )
    
    parser.add_argument(    "-w" , "--path-working"         
                            , nargs='?'
                            , const=1
                            , type=str
                            , default = "."                
                            , dest    = "workingDir"           
                            , help    = "Define the path to the output generation place" )
    
    parser.add_argument(    "-f" , "--parameter-file"  
                            , nargs='?'
                            , const=1
                            , type=str
                            , default = "SmkInput.json"                
                            , dest    = "parameterFile"    
                            , help    = "Define the log level " )
    
    parser.add_argument(    "--user-config"                      
                            , nargs='?'
                            , const=1
                            , type=str
                            , default = "SmkUserCfg.json" 
                            , dest    = "userConfigFile"   
                            , help    = "This file will overwrite configs from the input file!!!!" )
    
    parser.add_argument(    "--ll" , "--log-level"       
                            , nargs='?'
                            , const=1
                            , type=str
                            , default = None       
                            , dest    = "LogLevel"         
                            , help    = "Define the log level " )
    
    parser.add_argument(    "--lf" , "--log-file"        
                            , nargs='?'
                            , const=1
                            , type=str
                            , default = None       
                            , dest    = "LogFile"          
                            , help    = "Define the log file "  )

    return parser


def start_outlook():
    # Path to Outlook executable
    outlook_path = r"C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.EXE"
    # Start Outlook
    process = subprocess.Popen(outlook_path)
    return process

def kill_outlook(process):
    # Kill the Outlook process
    process.terminate()
    try:
        # Wait up to 10 seconds for the process to terminate
        print("wait")
        process.wait(timeout=10)  
    except subprocess.TimeoutExpired:
        # Forcefully kill if it doesn't terminate within the timeout
        print("process.kill")
        process.kill()  


def main():
    print("Start outlook")
    # Start Outlook
    outlook_process = start_outlook()
    # Allow some time for Outlook to process emails
    time.sleep(10)
    # Kill Outlook
    kill_outlook(outlook_process)
    parser = my_func_that_returns_a_parser()
    
    args = vars( parser.parse_args() )
    args['output'] = args[ 'workingDir' ]
    if not args[ 'LogLevel' ] == None :
        LOGGING_MAPPING = {
            "DEBUG"    : logging.DEBUG,
            "INFO"     : logging.INFO,
            "WARNING"  : logging.WARNING,
            "ERROR"    : logging.ERROR,
            "CRITICAL" : logging.CRITICAL
        }
        logging.basicConfig(
            level = LOGGING_MAPPING[ args[ 'LogLevel' ] ],
            format = "%(levelname)s {%(pathname)s:%(lineno)d} - %(message)s",
            handlers=[
                        # logging.FileHandler( args[ 'LogFile' ] ),
                        logging.StreamHandler()
                ]
        )
    
    logging.info( "Start Vector_IssueApp ... ")
    
    print("Start Vector_IssueApp ... ")
    tool = Vector_Issue.Vector_IssueApp()
    
    # run the tool
    #tool.main()
    
    logging.info( "End Vector_IssueApp ... ")
    


if __name__ == '__main__' :
    main()
    pass # Place for a breakpoint
    
    