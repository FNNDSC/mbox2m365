try:
    from        .               import __pkg, __version__
except:
    from __init__               import __pkg, __version__

# System imports
import      os, sys
sys.path.insert(1, os.path.join(os.path.dirname(__file__), '..'))
import      json
from        pathlib             import  Path

# Project specific imports
import      pfmisc
from        pfmisc._colors      import  Colors
from        pfmisc              import  other
from        pfmisc              import  error

import      pudb
import      mailbox
from        email.utils         import  make_msgid

import      pathlib

from        jobber              import  jobber
from        appdirs             import  *
import      shutil

import      time
import      base64
import      re

class Mbox2m365(object):
    """

    The core class for filtering a message out of an mbox and retransmitting
    it using m365

    """

    _dictErr = {
        'm365NotFound'      : {
            'action'        : 'checking on m365, ',
            'error'         : 'I could not find the executable in the current PATH.',
            'exitCode'      : 1},
        'm365'              : {
            'action'        : 'trying to use m365, ',
            'error'         : 'an error occured in using the app.',
            'exitCode'      : 2},
        'mbox'              : {
            'action'        : 'checking the mbox file, ',
            'error'         : 'the mbox seemed inaccessible. Does it exist?',
            'exitCode'      : 3},
        'outputDirFail'     : {
            'action'        : 'trying to check on the output directory, ',
            'error'         : 'directory not specified. This is a *required* input.',
            'exitCode'      : 4},
        'parseMessageIndices'     : {
            'action'        : 'trying to check on user specified message indices, ',
            'error'         : 'some error was triggered',
            'exitCode'      : 5},
        }

    def declare_selfvars(self):
        """
        A block to declare self variables -- this is a convenient place
        to simply "declare" all the self variables for reference in this
        object.
        """

        #
        # Object desc block
        #
        self.str_desc           : str   = self.args['desc']
        self.__name__           : str   = self.args['name']
        self.version            : str   = self.args
        self.tic_start          : float =  0.0
        self.verbosityLevel     : int   = -1
        self.d_m365             : dict  = {}
        self.configPath         : Path  = '/'
        self.keyParsedFile      : Path  = 'someFile.json'
        self.transmissionCmd    : Path  = 'someFile.rec'
        self.emailFile          : Path  = 'someFile.txt'
        self.l_keysParsed       : list  = []
        self.l_keysInMbox       : list  = []
        self.l_keysToParse      : list  = []
        self.l_keysTransmitted  : list  = []
        self.str_m365Path       : str   = ""
        self.dp                 = None
        self.log                = None
        self.mbox               = None
        self.o_msg              = None

    def __init__(self, *args, **kwargs):
        """
        Main constructor -- this method mostly defines variables declared
        in declare_selfvars()
        """
        self.args           = args[0]

        # The 'self' isn't fully instantiated, so
        # we call the following method on the class
        # directly.
        Mbox2m365.declare_selfvars(self)

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
        self.configPath     = Path(user_config_dir(self.__name__))
        self.keyParsedFile  = self.configPath / Path('keysParsed.json')
        self.transmissionCmd= self.configPath / Path('m365.cmd')
        self.mboxPath       = Path(self.args['inputDir']) / Path(self.args['mbox'])

    def env_check(self, *args, **kwargs) -> dict:
        """
        This method provides a common entry for any checks on the
        environment (input / output dirs, etc)
        """
        b_status    : bool  = True
        str_error   : str   = ''

        def mbox_check():
            # Check on the mbox
            self.log("Checking on mbox file...")
            try:
                self.mbox           = mailbox.mbox(str(self.mboxPath))
            except:
                self.log('mboxPath = %s' % str(self.mboxPath), comms = 'error')
                error.fatal(self, 'mbox', drawBox = True)
            self.l_keysInMbox       = self.mbox.keys()

        def configDir_check():
            # Check on config dir
            self.log("Checking on config dir '%s'..." % self.configPath)
            if not self.configPath.exists():
                self.configPath.mkdir(parents = True, exist_ok = True)

        def keysParsedFile_check():
            # Check on keysParsed file
            self.log("Checking on parsed history '%s'..." % self.keyParsedFile)
            if not self.keyParsedFile.exists():
                with self.keyParsedFile.open("w", encoding = "UTF-8") as f:
                    json.dump({"keysParsed":   []}, f)
                self.l_keysParsed   = []
            else:
                with self.keyParsedFile.open("r") as f:
                    self.l_keysParsed   = json.load(f)['keysParsed']

        def m365_checkOnPath():
            self.log("Checking for 'm365' executable on path...")
            self.str_m365Path = shutil.which('m365')
            if not self.str_m365Path:
                str_error   = 'm365 not found on path.'
                self.log(str_error, comms = 'error')
                error.fatal(self, 'm365NotFound', drawBox = True)
                b_status    = False

        mbox_check()
        configDir_check()
        keysParsedFile_check()
        m365_checkOnPath()

        return {
            'status':       b_status,
            'str_error':    str_error
        }

    def message_listToProcess(self):
        """
        Determine which new message in the mbox to process. Usually this is
        simply the difference between the existing keys in the mbox and the
        current keys.

        In the case when the user might have specified a specific set of
        indices with the ``--parseMessageIndices`` flag, process these
        keys instead.
        """
        if len(self.args['parseMessageIndices']):
            try:
                self.l_keysToParse  = self.args['parseMessageIndices'].split(',')
            except:
                error.fatal(self, 'parseMessageIndices', drawBox = True)
        else:
            self.l_keysToParse  = list(set(self.l_keysInMbox) - set(self.l_keysParsed))
        self.log('Message keys to transmit: %s' % self.l_keysToParse)
        return self.l_keysToParse

    def message_extract(self, index):
        """
        Simply extract the message (with very rudimentary error checking) from mbox
        """
        b_status: bool  = True
        try:
            self.o_msg  = self.mbox[self.mbox.keys()[index]]
        except:
            b_status    = False
        self.log("Extracted message at index '%s'" % index)

        return {
            'status':   b_status,
            'index':    index
        }

    def multipart_saveToFile(self, message):
        """
        Save the message to a file. In such a case m365 will be instructed
        to transmit the email from this file.
        """
        def urlify(s):
            # Remove all non-word characters (everything except numbers and letters)
            s = re.sub(r"[^\w\s]", '', s)

            # Replace all runs of whitespace with a single dash
            s = re.sub(r"\s+", '-', s)
            return s


        str_subjClean       : str   = urlify(self.d_m365['subject'])
        self.emailFile              = self.configPath / Path(str_subjClean + ".txt")
        with open(str(self.emailFile), "w") as f:
            f.write(message)

    def multipart_appendSimply(self, message):
        """
        Simple (naive) multipart handler. If sending a multipart message
        this attempts to base64 encode any attachments and append them
        into the message
        """
        def chunkstring(string, length):
            return (string[0+i:length+i] for i in range(0, len(string), length))

        def boundaryHeader_generate():
            nonlocal boundary, content_type, content_disposition, content_encoding
            return f"""
--------------{boundary}
Content-Type: {content_type}{filename}
Content-Disposition: {content_disposition}
Content-Transfer-Encoding: {content_encoding}

"""

        def boundaryFooter_generate():
            nonlocal boundary
            return f"""
--------------{boundary}--
"""

        str_body        : str   = ""
        bodyFirst               = "This is a multi-part message in MIME format."
        for part in message.walk():
            content_type        = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            content_encoding    = str(part.get("Content-Transfer-Encoding"))
            self.log("Processing multipart message...")
            self.log("Content-Type: " + f"{content_type}")
            self.log("Content-Disposition: " + f"{content_disposition}")
            self.log("Content-Transfer-Encoding: " + f"{content_encoding}")
            try:
                bodyPart = part.get_payload(decode = True).decode()
                self.log('bodyPart attached after decode()', comms = 'status')
            except:
                if self.args['b64_encode']:
                    bodyPart                = part.get_payload(decode = True)
                    if bodyPart:
                        boundary            = f'{make_msgid()}'
                        try:
                            filename        = f"; name={part.get_filename()}"
                        except:
                            filename        = ""
                        self.log('Encoding into base64 "ascii" for retransmission')
                        self.log('Original (<type>, <size>) = (%s, %s)' % \
                                    (type(bodyPart), len(bodyPart)))

                        bytes_b64   = base64.b64encode(bodyPart)
                        bodyPart    = bytes_b64.decode("ascii")
                        self.log('Encoded (<type>, <size>) = (%s, %s)' % \
                                    (type(bodyPart), len(bodyPart)))
                        length      = 72
                        bodyPartFixedWidth  = ""
                        for chunk in chunkstring(bodyPart, length):
                            bodyPartFixedWidth  += chunk + '\n'
                        bodyPart    = boundaryHeader_generate() + \
                                      bodyPartFixedWidth        + \
                                      boundaryFooter_generate()
                    else:
                        bodyPart    = ""
                else:
                    self.log('bodyPart attachment skipped!', comms = 'status')
                    bodyPart    = ""
            str_body += str(bodyPart)
        return bodyFirst + '\n' + str_body

    def message_parse(self, d_extract):
        """
        Populate the internal self.d_m365 payload
        """
        b_status        : bool  = False
        message                 = None
        if d_extract['status']:
            self.d_m365['subject']      = self.o_msg['Subject']
            self.d_m365['to']           = self.o_msg['Delivered-To']
            if self.o_msg.is_multipart():
                message                 = self.multipart_appendSimply(self.o_msg)
            else:
                message                 = self.o_msg.get_payload()
            if self.args['sendFromFile']:
                self.multipart_saveToFile(message)
                self.log('Email saved to %s' % self.emailFile)
            self.d_m365['bodyContents'] = message
            b_status    = True
            self.log(
                "Parsed message to '%s' re '%s'" % (
                                self.d_m365['to'],
                                self.d_m365['subject']
                        )
                )
        return {
            'status'    : b_status,
            'to'        : self.d_m365['to'],
            'subject'   : self.d_m365['subject'],
            'extract'   : d_extract
        }

    def message_transmit(self, d_parse):
        """
        Transmit the actual message -- either directly on the CLI or
        from a created file...
        """
        b_status        : bool  = False
        str_m365        : str   = ""
        d_m365          : dict  = {}
        # pudb.set_trace()
        if d_parse['status']:
            shell       = jobber.jobber({'verbosity': 1, 'noJobLogging': True})
            if self.args['sendFromFile']:
                self.d_m365['bodyContents'] = f'@{self.emailFile}'
            str_m365    = """#!/bin/bash

            m365 outlook mail send -s '%s' -t %s --bodyContents '%s'
            """ % \
                    (
                      self.d_m365['subject'],
                      self.d_m365['to'],
                      self.d_m365['bodyContents']
                    )
            with open(self.transmissionCmd, "w") as f:
                f.write(f'%s' % str_m365)
            self.transmissionCmd.chmod(0o755)
            d_m365      = shell.job_run(str(self.transmissionCmd))
            self.log(
                "Transmitted message, return code '%s'" % d_m365['returncode']
            )
            b_status    = True
        return {
            'status'    : b_status,
            'm365'      : d_m365,
            'parse'     : d_parse
        }

    def state_save(self):
        """
        Save the state of the system, i.e. save the list of parsed
        keys.
        """
        with self.keyParsedFile.open("w", encoding = "UTF-8") as f:
            json.dump({"keysParsed":   self.l_keysParsed}, f)

    def run(self, *args, **kwargs) -> dict:

        b_status            : bool  = False
        b_timerStart        : bool  = False
        d_env               : dict  = {}
        ld_send             : list  = []
        d_filter            : dict  = {}
        b_JSONprint         : bool  = True
        b_catchStragglers   : bool  = True
        d_env               : dict  = self.env_check()

        if d_env['status']:
            for k, v in kwargs.items():
                if k == 'timerStart':   b_timerStart    = bool(v)
                if k == 'JSONprint':    b_JSONprint     = bool(v)

            if b_timerStart:    other.tic()
            while b_catchStragglers:
                for message in self.message_listToProcess():
                    ld_send.append(
                        self.message_transmit(
                            self.message_parse(
                                self.message_extract(message)
                            )
                        )
                    )
                    self.l_keysTransmitted.append(message)
                self.l_keysParsed.extend(self.l_keysTransmitted)
                # Now check the environment again to grab any "stragglers"
                # i.e. emails that might have appeared *while* we checked the
                # first time
                self.state_save()
                self.log(
                    "Checking for stragglers that might have snuck in while we were processing...",
                    comms = 'tx'
                )
                d_env   = self.env_check()
                if not len(self.message_listToProcess()):
                    b_catchStragglers = False
                    self.log(
                        "No stragglers found... going back to sleep!",
                        comms = 'rx'
                    )
        d_ret           : dict = {
            'env'       : d_env,
            'runTime'   : other.toc(),
            'send'      : ld_send
        }

        return d_ret

