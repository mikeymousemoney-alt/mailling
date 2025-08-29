#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
================================
Vector_IssueCmd
================================

The Vector_IssueCmd program will do something. 


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
import os
import signal
import argparse
from pathlib import Path
import logging
import psutil


def my_func_that_returns_a_parser():
    # Load cammand line options
    parser = argparse.ArgumentParser(
                    prog = 'Vector_IssueCmd',
                    description = __doc__ ,
                    epilog = __copyright__ , 
                    formatter_class = argparse.RawTextHelpFormatter )

    parser.add_argument(     "--pl", "--plugins"
                            , nargs='+'
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
    
    parser.add_argument(    '--start-date'
                            , type=str
                            #, required= True
                            , dest    = "startDate"
                            , help    = 'Start date in the format DDMMYYYY')
                        

    return parser


def main():
    
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
    
    logging.info( "Start Vector_IssueCmd ... ")
    
    #print('Add some code')
    TemplatePath = Path(__file__).parent / '..' / 'templates'
    #print( 'Your template path is here: {}'.format(TemplatePath.resolve().as_posix() ) )

    valid_args = ['startDate']
    filtered_args = {k: v for k, v in args.items() if k in valid_args}

    tool = Vector_Issue.Vector_IssueApp( **filtered_args )
    
    # run the tool
    #tool.mainloop(  )

    logging.info( "End Vector_IssueCmd ... ")
    print( "End Vector_IssueCmd ... ")


if __name__ == '__main__' :
    main()
    pass # Place for a breakpoint
    
    