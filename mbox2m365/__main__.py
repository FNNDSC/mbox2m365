#!/usr/bin/env python3
#
# (c) 2022 Fetal-Neonatal Neuroimaging & Developmental Science Center
#                   Boston Children's Hospital
#
#              http://childrenshospital.org/FNNDSC/
#                        dev@babyMRI.org
#

import  sys, os
import  pudb

try:
    from    .               import mbox2m365
    from    .               import __pkg, __version__
except:
    from mbox2m365          import Mbox2m365
    from __init__           import __pkg, __version__
from    argparse            import RawTextHelpFormatter
from    argparse            import ArgumentParser

import  pfmisc
from    pfmisc._colors      import Colors
from    pfmisc              import other

import  json

str_desc = Colors.CYAN + r"""


                        _               ___            ____    __ _____
                       | |             |__ \          |___ \  / /| ____|
              _ __ ___ | |__   _____  __  ) |_ __ ___   __) |/ /_| |__
             | '_ ` _ \| '_ \ / _ \ \/ / / /| '_ ` _ \ |__ <| '_ \___ \
             | | | | | | |_) | (_) >  < / /_| | | | | |___) | (_) |__) |
             |_| |_| |_|_.__/ \___/_/\_\____|_| |_| |_|____/ \___/____/



                        (unix) m(ail)box 2 (MS Outlook)365

        This small python utility and module parses a standard (un*x) mailbox
        file for a message, and then transmits that message using an authent-
        icated m365 layer.

                             -- version """ + \
             Colors.YELLOW + __version__ + Colors.CYAN + """ --

        While useful in its own right as a standalone program, `mbox2m365` is
        most optimally applied as part of a properly configured `postfix` mail
        server. By assigning an (any) email client to use this `postfix` server
        and by monitoring changes in the mbox file, it is possible to allow


""" + Colors.NO_COLOUR

package_CLIcore = """
        [--inputDir <inputDir>]                                                 \\
        [--outputDir <outputDir>]                                               \\
        [--inputFile <inputFile>]                                               \\
        [--printElapsedTime]                                                    \\
        [--man]                                                                 \\
        [--synopsis]                                                            \\
        [--verbosity <verbosity>]                                               \\
        [--version]                                                             \\
        [--json]
"""

package_CLIself = '''
        --mbox <mbox>                                                           \\
        [--parseMessageIndices <M,N,...>]                                       \\
        [--b64_encode]                                                          \\
        [--sendFromFile]                                                        \\
        [--waitForStragglers <waitSecondsFirst>]                                \\
        [--cleanUp]                                                             \\'''

package_argSynopsisSelf = """
        --mbox <mbox>
        The mbox file to process.

        [--parseMessageIndices <M,N,...>]
        Instead of processing "new" messages in the mbox, explicitly
        filter and send the message at the specified key indices <M>, <N>...
        This is comma-separated string.

        [--b64_encode]
        If specified, encode any attachments with base64 encoding.

        [--sendFromFile]
        If specified, write body text to a file, and transmit the file. If not
        specified, the body text is added inline to the transmit command. For
        long messages, inline transmission will result in CLI overflow, but
        is fractionally faster for short messages.

        [--waitForStragglers <loopDelay>]
        This is a "hack" to delay processing a tad. For messages with many
        recipients, data might still be in the process of being appended to the
        mbox when this bridge opens. In that event, "late comers" might be
        missed.

        This is a naive switch that examines the --mbox filesize in a
        <loopDelay> waits until the filesize is unchanging between two
        successive loops. In this manner, the bridge attempts to wait until
        the mbox file size has stabilized.

        [--cleanUp]
        If specified, will clean up any files generated during operation.
"""

package_argSynopsisCore = """
        [--inputDir <inputDir>]
        Input base directory (default `/var/mail`)

        [--outputDir <outputDir>]
        The output root directory that will contain a tree structure identical
        to the input directory, and each "leaf" node will contain the analysis
        results.

        [--inputFile <inputFile>]
        A specific file in the <inputDir> to process.

        [--printElapsedTime]
        If specified, print the program execution time duration.

        [--man]
        Show full help.

        [--synopsis]
        Show brief help.

        [--json]
        If specified, output a JSON dump of final return.

        [--verbosity <level>]
        Set the app verbosity level.

            0: No internal output;
            1: Run start / stop output notification;
            2: As with level '1' but with simpleProgress bar in 'pftree';
            3: As with level '2' but with list of input dirs/files in 'pftree';
            5: As with level '3' but with explicit file logging for
                    - read
                    - analyze
                    - write
"""

def synopsis(ab_shortOnly = False):
    scriptName = os.path.basename(sys.argv[0])
    shortSynopsis =  """
    NAME

        mbox2m365

    SYNOPSIS

        mbox2m365                                                               \\""" + package_CLIself + package_CLIcore + """

    BRIEF EXAMPLE

        cd /var/mail
        find rudolph | entr mbox2m365 --inputFile rudolph
    """

    description =  '''
    DESCRIPTION

        ``mbox2m365`` allows for selective filtering of a message in an mbox
        format mail file and the re-transmission of that message using Outlook
        via the ``m365`` command line utility tool.

    ARGS ''' + package_argSynopsisCore + package_argSynopsisSelf + '''


    EXAMPLES

        cd /var/mail
        find rudolph | entr mbox2m365 --mbox rudolph                            \\
                                      --b64_encode --sendFromFile               \\
                                      --cleanUp --waitForStragglers 10

    '''

    if ab_shortOnly:
        return shortSynopsis
    else:
        return shortSynopsis + description

parserCore  = ArgumentParser(description        = 'Core I/O',
                             formatter_class    = RawTextHelpFormatter,
                             add_help           = False)
parserSelf  = ArgumentParser(description        = 'Self specific',
                             formatter_class    = RawTextHelpFormatter,
                             add_help           = False)

parserSelf.add_argument("--mbox",
                    help    = "the mbox file to analze",
                    dest    = 'mbox',
                    default = '')
parserCore.add_argument("--b64_encode",
                    help    = "encode all attachments as explicit inline base64",
                    dest    = 'b64_encode',
                    action  = 'store_true',
                    default = False)
parserCore.add_argument("--cleanUp",
                    help    = "clean up all operational files generated during execution",
                    dest    = 'cleanUp',
                    action  = 'store_true',
                    default = False)
parserCore.add_argument("--sendFromFile",
                    help    = "save the email to a file first, and then transmit this file",
                    dest    = 'sendFromFile',
                    action  = 'store_true',
                    default = False)
parserSelf.add_argument("--waitForStragglers",
                    help    = "wait buffer before starting processing",
                    dest    = 'waitForStragglers',
                    default = '1')
parserSelf.add_argument("--parseMessageIndices",
                    help    = "parse messages at given indices",
                    dest    = 'parseMessageIndices',
                    default = '')
parserSelf.add_argument("--printElapsedTime",
                    help    = "print program run time",
                    dest    = 'printElapsedTime',
                    action  = 'store_true',
                    default = False)

parserCore.add_argument("--inputDir",
                    help    = "input dir",
                    dest    = 'inputDir',
                    default = '/var/mail')
parserCore.add_argument("--inputFile",
                    help    = "input file",
                    dest    = 'inputFile',
                    default = '')
parserCore.add_argument("--outputDir",
                    help    = "output image directory",
                    dest    = 'outputDir',
                    default = '')
parserCore.add_argument("--man",
                    help    = "man",
                    dest    = 'man',
                    action  = 'store_true',
                    default = False)
parserCore.add_argument("--synopsis",
                    help    = "short synopsis",
                    dest    = 'synopsis',
                    action  = 'store_true',
                    default = False)
parserCore.add_argument("--json",
                    help    = "output final return in json",
                    dest    = 'json',
                    action  = 'store_true',
                    default = False)
parserCore.add_argument("--verbosity",
                    help    = "verbosity level for app",
                    dest    = 'verbosity',
                    default = "1")
parserCore.add_argument('--version',
                    help    = 'if specified, print version number',
                    dest    = 'b_version',
                    action  = 'store_true',
                    default = False)

def main(argv=None):
    parser  = ArgumentParser(description        = str_desc,
                             formatter_class    = RawTextHelpFormatter,
                             parents            = [parserCore, parserSelf])
    args = parser.parse_args()
    if args.man or args.synopsis:
        print(str_desc)
        if args.man:
            str_help     = synopsis(False)
        else:
            str_help     = synopsis(True)
        print(str_help)
        return 1

    if args.b_version:
        print("Name:    %s\nVersion: %s" % (__pkg.name, __version__))
        return 1

    args.version        = __version__
    args.name           = __pkg.name
    args.desc           = synopsis(True)

    try:
        mbox_2_m365         = Mbox2m365(vars(args))
    except:
        mbox_2_m365         = mbox2m365.Mbox2m365(vars(args))
    mbox_2_m365.log("%s:%s waking up!" % (__pkg.name, __version__), comms = 'tx')
    mbox_2_m365.log('Messaging object successfully created.')

    # And now run it!
    d_mbox2m365         = mbox_2_m365.run(timerStart = True)

    if args.printElapsedTime:
        mbox_2_m365.log(
                "Elapsed time = %f seconds" %
                d_mbox2m365['runTime']
        )

    mbox_2_m365.log("%s:%s all done and going to sleep!" % (__pkg.name, __version__), comms = 'rx')

    return 0

if __name__ == "__main__":
    sys.exit(main())
