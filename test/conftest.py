import pytest
import logging
import pathlib

helpText_Skip = '''Option to skip tests they depents on test parameters.
Some tests depents on the trace32 or on a CANoe instance but maybe you don't
have the full setup. In case of you want run the test and see what tests are
not able and which tests have errors you can give this option. So you will
see the tests are skiped.

[default] %(default)s

'''

def pytest_addoption(parser):
    parser.addoption(   '--t32Path'
                        , action='store'
                        , default= ''
                        , help=helpText_Skip)

def pytest_configure(config):
    """
    Allows plugins and conftest files to perform initial configuration.
    This hook is called for every plugin and initial conftest
    file after command line options have been parsed.
    """
    pass


@pytest.fixture(scope='session', autouse=True)
def logging_format(level_name: str='TEST_STEP', level_number: int=32):

    def log_test_step(self, message ='', RequirementId = 'N/A', TestId = 'N/A', Description = '', Condition = '', Result : str = ''  , *args, **kwargs):

        if self.isEnabledFor(level_number):
            message = "\n\t".join( [ message 
            , "Requirement ID:\t"+  "\n\t\t\t".join( RequirementId.split('\n') )
            , "Test ID:\t"       +  "\n\t\t\t".join( TestId.split('\n') )
            , "Description:\t"   +  "\n\t\t\t".join( Description.split('\n') )
            , "Contition:\t"     +  "\n\t\t\t".join( Condition.split('\n') )
            , "Result:\t\t"      +  "\n\t\t\t".join( Result.split('\n') )
            ])      
            self._log(level_number, message, args, **kwargs)

    logging.addLevelName(level_number, level_name)
    setattr(logging, level_name, level_number)
    setattr(logging.getLoggerClass(), level_name.lower(), log_test_step)

@pytest.hookimpl()
def pytest_sessionstart( session):
    """
    Called after the Session object has been created and
    before performing collection and entering the run test loop.
    """
    logging.getLogger().info(" st_sessionstart")
    

def pytest_sessionfinish(session, exitstatus):
    """
    Called after whole test run finished, right before
    returning the exit status to the system.
    """
    logging.getLogger().info( "sessionfinish" )

def pytest_unconfigure(config):
    """
    called before test process is exited.
    """  
    logging.getLogger().info("pytest_unconfigure")
    