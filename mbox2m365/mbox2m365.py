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
import      hashlib

from        collections         import defaultdict

def field_get(message:dict, field:str) -> str:
    ret:str     = ""
    if 'Delivered-To' in field:
        if message.get(field):
            ret = str(message.get(field))
        elif message.get('To'):
            ret = str(message.get('To'))
    else:
            ret = str(message.get(field))
    return ret

def fields_get(message:dict, l_field:list, sep:str = "") -> str:
    ret:str     = ""
    val:str     = ""
    for field in l_field:
        val     = field_get(message, field)
        if val and val != 'None':
            if ret:
                ret += f"{sep}{val}"
            else:
                ret += val
    return ret

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
        self.configPath         : Path  = Path('/')
        self.keyParsedFile      : Path  = Path('someFile.json')
        self.transmissionCmd    : Path  = Path('someFile.cmd')
        self.emailFile          : Path  = Path('someFile.txt')
        self.l_attachments      : list  = []
        self.l_keysParsed       : list  = []
        self.l_keysInMbox       : list  = []
        self.l_keysToParse      : list  = []
        self.l_keysTransmitted  : list  = []
        self.str_m365Path       : str   = ""
        self.dp                 = None
        self.log                = None
        self.mbox               = None
        self.mboxPath           : Path  = Path('mbox')
        self.lo_msg             : list  = [] # list of objects of messages
        self.d_m365             : dict  = {}
        self.ld_m365            : list  = []

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
            'output':           'json',     # json,text,csv
            'bodyHash':         '',
            'attachments':      ()
        }

        self.dp             = pfmisc.debug(
                                 verbosity   = int(self.args['verbosity']),
                                 within      = self.__name__,
                                 methodcol   = 55
                             )
        self.log            = self.dp.qprint
        self.configPath     = Path(user_config_dir(self.__name__))
        self.keyParsedFile  = self.configPath / Path('keysParsed.json')
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

    def message_listToProcess(self) -> list:
        """
        Determine which new message(s) in the mbox to process. Usually this is
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
        self.log('Message keys to process: %s' % self.l_keysToParse)
        return self.l_keysToParse

    def messageList_extract(self, l_index : list) -> dict:
        """
        Simply extract the newly arrived mbox messages into an internal
        list buffer (with very rudimentary error checking)
        """
        lb_status   : bool  = []
        l_extracted : list  = []
        for index in l_index:
            try:
                self.lo_msg.append(self.mbox[self.mbox.keys()[index]])
                self.log("Extracted message at index '%s'" % index)
                lb_status.append(True)
                l_extracted.append(index)
            except:
                lb_status.append(False)

        return {
            'status'        :   any(lb_status),
            'l_index'       :   l_index,
            'l_extracted'   :   l_extracted
        }

    def urlify(self, s : str) -> str:
        """Simple method that removes all whitespace and non-word chars from
        a string -- typically to create a "clean" version suitable for simple
        processing

        Args:
            s (str): the string to process

        Returns:
            str: the clean version of the input string
        """
        # Remove all non-word characters (everything except numbers and letters)
        s = re.sub(r"[^\w\s]", '', s)
        # Replace all runs of whitespace with a single dash
        s = re.sub(r"\s+", '-', s)
        return s

    def body_saveToFile(self, m365message):
        """
        Save the message body to a file, allowing the transmission of messages
        that will otherwise exceed CLI string length.

        In such a case m365 will be instructed to transmit the email from this file.
        """

        with open(str(self.emailFile), "w") as f:
            f.write(m365message['bodyContents'])

    def multipart_separateAttachments(self, message: bytes) -> dict:
        """Separate attachments in the <message> into discrete files

        Args:
            message (bytes): mbox message to process

        Returns:
            dict: names of all detached attachment files
        """
        d_ret   = {
            'status'            : False,
            'textBody'          : '',
            'attachFilesList'   : []
        }
        str_body        : str   = ""
        for part in message.walk():
            content_type        = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            content_encoding    = str(part.get("Content-Transfer-Encoding"))
            self.log("\t\tProcessing multipart message...")
            self.log("\t\tContent-Type: " + f"{content_type}")
            self.log("\t\tContent-Disposition: " + f"{content_disposition}")
            self.log("\t\tContent-Transfer-Encoding: " + f"{content_encoding}")
            try:
                # If this decodes, we are in string part of the message
                bodyPart    = part.get_payload(decode = True).decode()
                d_ret['status'] = True
                self.log('\t\tbodyPart attached after decode()', comms = 'status')
                str_body   += str(bodyPart)
                d_ret['textBody']   = str_body
            except:
                # If this decodes, we are in a binary (attachment) part of
                # the message -- decode, save, and record in list
                bodyPart    = part.get_payload(decode = True)
                try:
                    # Determine the filename from the attachment
                    str_filename:str = eval(content_disposition.split('filename=')[1])
                    d_ret['attachFilesList'].append(str_filename)
                    # and save to self.configPath dir:
                    path_save:Path   = self.configPath / Path(str_filename)
                    with open(str(path_save), "wb") as bf:
                        bf.write(bodyPart)
                except:
                    pass
        return d_ret

    def multipart_appendSimply(self, message) -> str:
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
        bodyFirst               = "This is a multi-part message in MIME format." + '\n'
        for part in message.walk():
            content_type        = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            content_encoding    = str(part.get("Content-Transfer-Encoding"))
            self.log("\t\tProcessing multipart message...")
            self.log("\t\tContent-Type: " + f"{content_type}")
            self.log("\t\tContent-Disposition: " + f"{content_disposition}")
            self.log("\t\tContent-Transfer-Encoding: " + f"{content_encoding}")
            try:
                bodyPart = part.get_payload(decode = True).decode()
                self.log('\t\tbodyPart attached after decode()', comms = 'status')
            except:
                if self.args['b64_encode']:
                    bodyPart                = part.get_payload(decode = True)
                    if bodyPart:
                        boundary            = f'{make_msgid()}'
                        try:
                            filename        = f"; name={part.get_filename()}"
                        except:
                            filename        = ""
                        self.log('\t\tEncoding into base64 "ascii" for retransmission')
                        self.log('\t\t\tOriginal (<type>, <size>) = (%s, %s)' % \
                                    (type(bodyPart), len(bodyPart)))

                        bytes_b64   = base64.b64encode(bodyPart)
                        bodyPart    = bytes_b64.decode("ascii")
                        self.log('\t\t\tEncoded (<type>, <size>) = (%s, %s)' % \
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
                    self.log('\t\tbodyPart attachment skipped!', comms = 'status')
                    bodyPart    = ""
            str_body += str(bodyPart)
        return bodyFirst + '\n' + str_body

    def messageList_componentsSeparate(self, d_extract : dict) -> dict:
        """Separate the extracted message list into constituent 'subject',
        'body' and 'to' list components. Essentially creating lists of

            [<to1>, ... <toN>]
            [<subj1>, ... <subjN>]
            [<date1>, ... <dateN>]
            [<bodyContents1>, ... <bodyContentsN>]
            [<hashContents1>, ... <hashContentsN>]

        Args:
            d_extract (Dict): a dictionary containing the list of extracted
                              indices. Used mainly to determine the status
                              of the extract call.

        Returns:
            dict: a structure containing parsed 'subject', 'body', and 'to'
                as explicit lists.
        """
        b_status            : bool  = False
        message                     = None
        lstr_subject        : list  = []
        lstr_to             : list  = []
        lstr_date           : list  = []
        lstr_msgBody        : list  = []
        lhash_msgBody       : list  = []
        lstr_attachments    : list  = []
        count               : int   = 0

        if d_extract['status']:
            if self.lo_msg: b_status = True
            for message in self.lo_msg:
                # lstr_subject.append(message['Subject'])
                # lstr_to.append(message['Delivered-To'])
                # lstr_date.append(message['Date'])
                lstr_subject.append(field_get(message, 'Subject'))
                lstr_to.append(field_get(message, 'Delivered-To'))
                lstr_date.append(field_get(message, 'Date'))
                if message.is_multipart():
                    d_parts:dict = self.multipart_separateAttachments(message)
                    lstr_msgBody.append(d_parts['textBody'])
                    lstr_attachments.append(tuple(d_parts['attachFilesList']))
                else:
                    lstr_msgBody.append(message.get_payload())
                lhash_msgBody.append(hashlib.md5(lstr_msgBody[-1].encode('utf-8')).hexdigest())
                self.log(
                    "Parsed message [%03d-%s] on '%31s' to '%s' re '%s'" % (
                                    self.l_keysToParse[count],
                                    lhash_msgBody[-1],
                                    lstr_date[-1],
                                    lstr_to[-1],
                                    lstr_subject[-1]
                            )
                    )
                count += 1
        return {
            'status'    : b_status,
            'd_fields'  : {
                'l_date'        : lstr_date,
                'l_to'          : lstr_to,
                'l_subject'     : lstr_subject,
                'l_body'        : lstr_msgBody,
                'l_attachments' : lstr_attachments,
                'l_hash'        : lhash_msgBody
            },
            'prior'     : d_extract,
        }

    def messageList_tallyOccurences(self, d_separate : dict) -> dict:
        """Tally the occurrences of elements.

        Args:
            d_separate (dict): The various upstream separate lists

        Returns:
            dict: A dictionary of occurences of each field. This is
                  used downstream to collapse messages with identical
                  content into one transmission from Outlook.
        """
        b_status : bool = True

        def tally_occurences(seq):
            tally = defaultdict(list)
            for i,item in enumerate(seq):
                try:
                    tally[item].append(i)
                except:
                    # We're trying to tally a possible attachment list so can skip
                    pass
            return ({key: locs} for key,locs in tally.items()
                                    if len(locs)>=1)

        d_occurences    : dict = {
            'l_to'          : [],
            'l_date'        : [],
            'l_subject'     : [],
            'l_attachments' : [],
            'l_body'        : [],
            'l_hash'        : []
        }
        if d_separate['status']:
            for key in d_separate['d_fields'].keys():
                for occur in (tally_occurences(d_separate['d_fields'][key])):
                    d_occurences[key].append(occur)
        return {
            'status'        :   b_status,
            'd_occurences'  :   d_occurences,
            'prior'         :   d_separate
        }

    def messageList_consolidateOccurences(self, d_tally : dict) -> dict:
        """Collapse messages that are 'cc' or 'bcc' into one message
        within the internal list of message dictionaries.

        Args:
            d_occurences (dict): the tally of duplicate field occurences

        Returns:
            dict: dictionary of consolidated messages to transmit
        """

        def get(atype : str, what : str, l : str, i : int):
            """A quick nested helper to extract the key or value from
               a list 'l' at index 'i'

            Args:
                atype (str): Either a 'set' or a 'list' to get
                what (str): The 'key' or 'value' in the dictionary to extract
                l (str): The name of the d_occurences list to process
                i (int): The index within the list

            Returns:
                str: The key or value
            """
            target  = None
            if what == 'key':
                target = list(d_tally['d_occurences'][l][i].keys())[0]
            else:
                target = list(d_tally['d_occurences'][l][i].values())[0]
            if atype == 'set':
                return set(target)
            else:
                return target

        def lists_haveEqualValues(l_lists: list, idx: int) -> bool:
            """Check if variable number of lists are equal.

            Args:
                l_lists (list): The lists within the tally/occurences to process
                idx (int): The list index within the occurence dictionary

            Returns:
                bool: if the lists are equal in value, return True; else False
            """
            b_equal     = True
            l_set       = []
            for el in l_lists:
                l_set.append(get('set', 'value', el, idx))
            n = set.intersection(*l_set)
            for el in l_lists:
                if n != get('set', 'value', el, idx):
                    b_equal = False
            return b_equal

        def recipient_sanitize(address:str) -> list[str]:
            email_pattern = r'<([^>]+)>|([^\s,]+@[^\s,]+)'
            matches = re.findall(email_pattern, address)
            email_addresses = [
                                match[0] if match[0]
                                    else match[1] for match in matches if match[0] or match[1]
                            ]
            return email_addresses

        def recipients_get(l_target : list) -> str:
            """For each o_msg index in l_target, create a compound
            comma separated list of recipients

            Args:
                l_target (list): The list of indices in the original mbox
                                 message structure to consult for "to"
                                 recipients.

            Returns:
                str: The (comma separated) string (list) of recipients
            """
            str_ret:str     = ""
            l_to:list       = []
            for msgi in l_target:
                #pudb.set_trace()
                l_to.extend(recipient_sanitize(
                                fields_get(self.lo_msg[msgi],
                                          ['Delivered-To', 'Cc', 'Bcc'], ","))
                            )
                self.l_keysParsed.append(self.l_keysToParse[msgi])
            str_ret         = ','.join(l_to)
            return str_ret

        def attachments_get(l_messageList: list) -> tuple:
            """Resolve the attachments (if any) for a given message

            Args:
                l_messageList (list): a list of unique consolidated
                message indices

            Returns:
                tuple: tuple of string attachment filenames (w/o leading dir)
                       for this message
            """
            t_attachmentsForMessage:tuple   = ()
            l_attachmentGroup:list          = d_tally['d_occurences']['l_attachments']
            for group in l_attachmentGroup:
                if len(l_messageList):
                    if l_messageList[0] in list(group.values())[0]:
                        t_attachmentsForMessage = list(group.keys())[0]
            return t_attachmentsForMessage

        b_status    : bool  = False
        msgCount    : int   = 0

        if d_tally['status']:
            msgCount                = len(d_tally['d_occurences']['l_date'])
            if msgCount: b_status   = True
            for idx in range(0, msgCount):
                # This test was originally predicated on the idea that tallying
                # occurences across dates and hashes would always match. Sometimes
                # with attachments, however, it seems the hashes do not. Hence for
                # now this "dummy" fall through based solely on the date stamp
                # only.
                if lists_haveEqualValues(['l_date', 'l_date'], idx):
                    self.ld_m365.append(self.d_m365.copy())
                    # If there are less "subjects" than messages
                    # (N messages all with the same subject, use
                    # the first subject)
                    try:
                        self.ld_m365[idx]['subject']      = get('list', 'key', 'l_subject', idx)
                    except:
                        self.ld_m365[idx]['subject']      = get('list', 'key', 'l_subject', 0)
                    self.ld_m365[idx]['to']           = recipients_get(get('list', 'value', 'l_date', idx))
                    self.ld_m365[idx]['attachments']  = attachments_get(get('list', 'value', 'l_date', idx))
                    self.ld_m365[idx]['bodyContents'] = get('list', 'key', 'l_body', idx)
                    self.ld_m365[idx]['bodyHash']     = get('list', 'key', 'l_hash', idx)

        return {
            'status'        :   b_status,
            'messageCount'  :   msgCount,
            'ld_transmit'   :   self.ld_m365,
            'mboxIndices'   :   self.l_keysParsed,
            'prior'         :   d_tally
        }

    def messageList_transmit(self, d_consolidated : dict) -> dict:
        """Transmit the consolidated message list

        Args:
            d_consolidated (dict): a dictionary containing the list of
                                   messages to transmit

        Returns:
            dict: a dictionary of the transmission list record
        """

        def cleanAttachments(m365message) -> None:
            if not self.args['cleanUp']: return
            for attachment in m365message['attachments']:
                try:
                    (self.configPath / Path(attachment)).unlink()
                    self.log("\tRemoving attachment file '%s'" % attachment)
                except:
                    pass

        def cleanUp() -> None:
            if not self.args['cleanUp']: return
            self.log("\tRemoving tx file '%s'" % self.transmissionCmd)
            self.transmissionCmd.unlink()
            if self.args['sendFromFile']:
                self.log("\tRemoving body file '%s'" % self.emailFile)
                self.emailFile.unlink()

        def bodyToFile_check(m365message) -> None:
            if self.args['sendFromFile']:
                self.emailFile  = self.configPath / Path(baseFileName + "_body.txt")
                self.body_saveToFile(m365message)
                m365message['bodyContents']     = f'@{self.emailFile}'

        def txscript_content(m365message) -> str:
            str_m365    : str   = ""
            str_m365    = """#!/bin/bash

            m365 outlook mail send -s '%s' -t '%s' --bodyContents '%s'""" % \
                    (
                      m365message['subject'],
                      m365message['to'],
                      m365message['bodyContents']
                    )
            str_m365 = ''.join(str_m365.split(r'\r'))
            if len(m365message['attachments']):
                for attachment in m365message['attachments']:
                    str_m365 += ' --attachment "' + str(self.configPath / Path(attachment)) + '"'
            return str_m365

        def txscript_save(str_content) -> None:
            self.transmissionCmd = self.configPath / Path(baseFileName + "_tx.cmd")
            with open(self.transmissionCmd, "w") as f:
                f.write(f'%s' % str_content)
            self.transmissionCmd.chmod(0o755)

        b_status        : bool  = False
        d_m365          : dict  = {}
        ld_m365         : list  = []
        baseFileName    : str   = ""

        if d_consolidated['status']:
            shell       = jobber.jobber({'verbosity': 1, 'noJobLogging': True})
            # First send all the messages...
            for m365message in d_consolidated['ld_transmit']:
                b_status        = True
                baseFileName    = self.urlify(m365message['subject'])
                bodyToFile_check(m365message)
                txscript_save(txscript_content(m365message))
                d_m365  = shell.job_run(str(self.transmissionCmd))
                self.log(
                    "Transmitted message (%s), return code '%s', recipients '%s'" %\
                         (m365message['bodyHash'],
                          d_m365['returncode'],
                          m365message['to'])
                )
                b_status    = True
                ld_m365.append(d_m365.copy())
                cleanUp()
            for m365message in d_consolidated['ld_transmit']:
                cleanAttachments(m365message)

        return {
            'status'    : b_status,
            'ld_m365'   : ld_m365,
            'prior'     : d_consolidated
        }

    def state_save(self):
        """
        Save the state of the system, i.e. save the list of parsed
        keys.
        """
        with self.keyParsedFile.open("w", encoding = "UTF-8") as f:
            json.dump({"keysParsed":   self.l_keysParsed}, f)

    def mbox_pauseUntilSizeStable(self) -> dict:
        """A simple check that examines the size of the mbox file
        in a tight loop and waits until the size is unchanged over
        two successive loops before continuing.
        """

        d_env           : dict  = {}
        b_sizeStable    : bool  = False
        mboxSizeOld     : Path  = Path(self.mboxPath).stat().st_size
        mboxSizeNew     : Path  = Path(self.mboxPath).stat().st_size
        while not b_sizeStable:
            self.log('mbox file size currently: %d' % mboxSizeOld, comms = 'status')
            self.log('waiting %s seconds for stragglers...' % \
                        self.args['waitForStragglers'])
            time.sleep(int(self.args['waitForStragglers']))
            mboxSizeNew = Path(self.mboxPath).stat().st_size
            if mboxSizeNew == mboxSizeOld:
                b_sizeStable    = True
                self.log('mbox size stable, continuing with processing...')
                d_env           = self.env_check()
            else:
                mboxSizeOld     = mboxSizeNew
                self.log('mbox size unstable, waiting...')
        return d_env

    def run(self, *args, **kwargs) -> dict:

        b_status            : bool  = False
        b_timerStart        : bool  = False
        d_env               : dict  = {}
        ld_send             : list  = []
        d_filter            : dict  = {}
        b_catchStragglers   : bool  = True
        d_env               : dict  = self.env_check()
        d_tx                : dict  = {}

        if d_env['status']:
            for k, v in kwargs.items():
                if k == 'timerStart':   b_timerStart    = bool(v)
                if k == 'JSONprint':    b_JSONprint     = bool(v)

            if b_timerStart:    other.tic()

            self.mbox_pauseUntilSizeStable()

            while self.message_listToProcess():
                d_tx =  self.messageList_transmit(
                            self.messageList_consolidateOccurences(
                                self.messageList_tallyOccurences(
                                    self.messageList_componentsSeparate(
                                        self.messageList_extract(
                                            self.message_listToProcess()
                                        )
                                    )
                                )
                            )
                        )
                self.state_save()

        d_ret           : dict = {
            'env'       : d_env,
            'runTime'   : other.toc(),
            'send'      : d_tx
        }

        return d_ret

