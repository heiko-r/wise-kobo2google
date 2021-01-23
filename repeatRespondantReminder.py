# Make sure that the locale 'en_SG.UTF-8' is installed!

from common import *

'''
Returns required configuration file name list.
'''
def getConfigFileList():
    return [GOOGLE_TOKENS_FILE_NAME]

'''
Perform environment checks to ensure all required settings are present.
'''
def _checkEnvironment():
    # Check all required configuration files present and connected to internet.
    checkEnvironment(getConfigFileList(), isInternetCheckRequired=True)

################################################################################

'''
Main function
'''
def main(argv):
    # Check environment
    _checkEnvironment()

    # todo: implementation

if __name__ == '__main__':
    main(sys.argv[1:])
