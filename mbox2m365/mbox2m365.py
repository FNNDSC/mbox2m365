from        .                   import __pkg, __version__

# System imports
import      os
import      json
from        pathlib             import  Path

# Project specific imports
import      pfmisc
from        pfmisc._colors      import  Colors
from        pfmisc              import  other
from        pfmisc              import  error

import      pudb
import      mailbox

import      pathlib

import      jobber

class Mbox2m365(object):
    """

    The core class for filtering a message out of an mbox and retransmitting
    it using m365

    """

    _dictErr = {
        'm365'   : {
            'action'        : 'trying to use m365, ',
            'error'         : 'an error occured in using the app.',
            'exitCode'      : 1},
        'mbox'   : {
            'action'        : 'checking the mbox file, ',
            'error'         : 'the mbox seemed inaccessible. Does it exist?',
            'exitCode'      : 2},
        }

    def declare_selfvars(self):
        """
        A block to declare self variables
        """

        #
        # Object desc block
        #
        self.str_desc                   = self.args['desc']
        self.__name__                   = self.args['name']
        self.version                    = self.args['version']

        self.dp                         = None
        self.log                        = None
        self.tic_start                  = 0.0
        self.verbosityLevel             = -1

    def __init__(self, *args, **kwargs):
        """
        Main constructor
        """
        self.args           = args[0]

        # The 'self' isn't fully instantiated, so
        # we call the following method on the class
        # directly.
        Mbox2m365.declare_selfvars(self)

        self.indexToParse   = self.args['parseMessageIndex']
        self.fromOverride   = self.args['fromOverride']

        self.d_m365         = {
            'subject':          '',
            'to':               '',
            'bodyContents':     '',
            'bodyContentType':  'Text',     # Text,HTML
            'saveToSentItems':  'false',    # false,true
            'output':           'json'      # json,text,csv
        }

        self.dp             = pfmisc.debug(
                                 verbosity   = int(self.args['verbosity']),
                                 within      = self.__name__
                             )
        self.log            = self.dp.qprint
        self.mboxPath       = Path(self.args['inputDir']) / Path(self.args['inputFile'])
        try:
            self.mbox           = mailbox.mbox(str(self.mboxPath))
        except:
            self.log('mboxPath = %s' % str(self.mboxPath), comms = 'error')
            error.fatal(self, 'mbox', drawBox = True)
        self.o_msg          = None

    def message_extract(self):
        """
        Simply extract the message (with error checking) from mbox
        """
        b_status: bool  = True
        try:
            self.o_msg  = self.mbox[self.mbox.keys()[int(self.args['parseMessageIndex'])]]
        except:
            b_status    = False

        return {
            'status':   b_status
        }

    def message_parse(self, d_extract):
        """
        Populate the internal self.d_m365 payload
        """
        b_status        : bool  = False
        if d_extract['status']:
            self.d_m365['subject']      = self.o_msg['Subject']
            self.d_m365['to']           = self.o_msg['To']
            self.d_m365['bodyContents'] = self.o_msg.get_payload()
            b_status    = True
        return {
            'status':   b_status
        }

    def message_transmit(self, d_parse):
        """
        Transmit the actual message
        """
        b_status        : bool  = False
        d_m365          : dict  = {}
        if d_parse['status']:
            shell       = jobber.jobber({'verbosity': 1, 'noJobLogging': True})
            str_m365    : str   = "m365 outlook sendmail -s '%s' -t %s --bodyContents '%s'" % \
                                  (
                                    self.d_m365['subject'],
                                    self.d_m365['to'],
                                    self.d_m365['bodyContents']
                                  )
            d_response  : dict  = shell.job_run(str_m365)
            b_status    = True
        return {
            'status'    : b_status,
            'm365'      : d_m365
        }

    def env_check(self, *args, **kwargs) -> dict:
        """
        This method provides a common entry for any checks on the
        environment (input / output dirs, etc)
        """
        b_status    : bool  = True
        str_error   : str   = ''

        if not len(self.args['outputDir']):
            b_status = False
            str_error   = 'output directory not specified.'
            self.dp.qprint(str_error, comms = 'error')
            error.warn(self, 'outputDirFail', drawBox = True)

        return {
            'status':       b_status,
            'str_error':    str_error
        }

    def run(self, *args, **kwargs) -> dict:

        b_status        : bool  = False
        b_timerStart    : bool  = False
        d_env           : dict  = {}
        d_filter        : dict  = {}
        b_JSONprint     : bool  = True

        for k, v in kwargs.items():
            if k == 'timerStart':   b_timerStart    = bool(v)
            if k == 'JSONprint':    b_JSONprint     = bool(v)

        if b_timerStart:    other.tic()
        pudb.set_trace()
        d_send = self.message_transmit(
                    self.message_parse(
                        self.message_extract()
                    )
                )

        d_ret           : dict = {
            'runTime'   :  other.toc(),
            'send'      : d_send
        }

        return d_ret

