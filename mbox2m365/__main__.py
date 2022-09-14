#!/usr/bin/env python3
#
# (c) 2022 Fetal-Neonatal Neuroimaging & Developmental Science Center
#                   Boston Children's Hospital
#
#              http://childrenshospital.org/FNNDSC/
#                        dev@babyMRI.org
#

import sys, os

try:
    from    .               import pfdo
    from    .               import __pkg, __version__
except:
    from pfdo               import pfdo
    from __init__           import __pkg, __version__
from    argparse            import RawTextHelpFormatter
from    argparse            import ArgumentParser
import  pudb

import  pfmisc
from    pfmisc._colors      import Colors
from    pfmisc              import other

str_desc = Colors.CYAN + """

                                  __      _
                                 / _|    | |
                          _ __  | |_   __| |  ___
                         | '_ \ |  _| / _` | / _ \
                         | |_) || |  | (_| || (_) |
                         | .__/ |_|   \__,_| \___/
                         | |
                         |_|



                          Path-File Do (something)

        Recursively walk down a directory tree and perform some operation
        on files in each directory (optionally filtered by some simple
        expression). Results of each operation are saved in output tree
        that  preserves the input directory structure.


                             -- version """ + \
             Colors.YELLOW + __version__ + Colors.CYAN + """ --

        'pfdo' is the base infrastructure app/class for walking down some
        dir tree, optionally finding and tagging files that conform to
        a simple filter, and running (something) on the files. This class
        provides the basic mechanism for using pfree in this manner and the
        (something) operation should be provided by a subclass of this base.

        As part of the "pf*" suite of applications, it is geared to IO as
        directories. Nested directory trees within some input directory
        are reconstructed in an output directory, preserving directory
        structure.


""" + Colors.NO_COLOUR

package_CLI = """
         -I|--inputDir <inputDir>                           \\
         -O|--outputDir <outputDir>                         \\
        [-i|--inputFile <inputFile>]                        \\
        [-f|--fileFilter <filter1,filter2,...>]             \\
        [-L|--filteFilterLogic AND|OR]                      \\
        [-d|--dirFilter <filter1,filter2,...>]              \\
        [--outputLeafDir <outputLeafDirFormat>]             \\
        [--printElapsedTime]                                \\
        [--threads <numThreads>]                            \\
        [--test]                                            \\
        [-x|--man]                                          \\
        [-y|--synopsis]                                     \\
        [--overwrite]                                       \\
        [--followLinks]                                     \\
        [--verbosity <verbosity>]                           \\
        [--version]                                         \\
        [--json]
"""

package_argSynopsis = """
        -I|--inputDir <inputDir>
        Input base directory to traverse.

        -O|--outputDir <outputDir>
        The output root directory that will contain a tree structure identical
        to the input directory, and each "leaf" node will contain the analysis
        results.

        [-i|--inputFile <inputFile>]
        An optional <inputFile> specified relative to the <inputDir>. If
        specified, then do not perform a directory walk, but convert only
        this file.

        [-f|--fileFilter <someFilter1,someFilter2,...>]
        An optional comma-delimated string to filter out files of interest
        from the <inputDir> tree. Each token in the expression is applied in
        turn over the space of files in a directory location according to a
        logical operation, and only files that contain this token string in
        their filename are preserved.

        [-L|--filteFilterLogic AND|OR]
        The logical operator to apply across the fileFilter operation. Default
        is OR.

        [-d|--dirFilter <someFilter1,someFilter2,...>]
        An additional filter that will further limit any files to process to
        only those files that exist in leaf directory nodes that have some
        substring of each of the comma separated <someFilter> in their
        directory name.

        [--outputLeafDir <outputLeafDirFormat>]
        If specified, will apply the <outputLeafDirFormat> to the output
        directories containing data. This is useful to blanket describe
        final output directories with some descriptive text, such as
        'anon' or 'preview'.

        This is a formatting spec, so

            --outputLeafDir 'preview-%s'

        where %s is the original leaf directory node, will prefix each
        final directory containing output with the text 'preview-' which
        can be useful in describing some features of the output set.

        [--test]
        If specified, run the "dummy" internal callback loop triad. The test
        flow simply tags files in some inputDir tree and "touches" them to a
        reconstiuted tree in the output directory, prefixed with the text
        "analyzed-".

        [--threads <numThreads>]
        If specified, break the innermost analysis loop into <numThreads>
        threads.

        [-x|--man]
        Show full help.

        [-y|--synopsis]
        Show brief help.

        [--json]
        If specified, output a JSON dump of final return.

        [--followLinks]
        If specified, follow symbolic links.

        [-v|--verbosity <level>]
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

        pfdo                                                \\

    SYNOPSIS

        pfdo """ + package_CLI + """

    BRIEF EXAMPLE

        pfdo                                                \\
            -I /var/www/html/data -f jpg                    \\
            -O /var/www/html/jpg                            \\
            --threads 0 --printElapsedTime
    """

    description =  '''
    DESCRIPTION

        ``pfdo`` provides the base mechanism for navigating some arbitrary
        tree, providing the base hooks for operating on (possibly filtered)
        files in each directory, and saving results in an output tree that
        reflects the input tree topology.

    ARGS ''' + package_argSynopsis + '''


    EXAMPLES

    Perform a `pfdo` down some input directory:

        pfdo                                                \\
            -I /var/www/html/data -f jpg                    \\
            -O /tmp/jpg --test --json                       \\
            --threads 0 --printElapsedTime

    The above will find all files in the tree structure rooted at
    /var/www/html/data that also contain the string "jpg" anywhere
    in the filename. For each file found, a corresponding file will
    be touched in the output directory, in the same tree location as
    the original input. This touched file will be prefixed with the
    string "analyzed-".

        pfdo                                                \\
            -I $(pwd)/raw  -d 100307 -f " "                 \\
            -O $(pwd)/out --test --json                     \\
            --threads 0 --printElapsedTime

    This will consider each directory in the input tree space that
    contains files, but will "tag" any leaf node directory that
    contains the string "100307" with a tag "file" "%d-100307".

    Finally the elapsed time and a JSON output are printed.

    '''

    if ab_shortOnly:
        return shortSynopsis
    else:
        return shortSynopsis + description

parser  = ArgumentParser(description = str_desc, formatter_class = RawTextHelpFormatter)

parser.add_argument("-I", "--inputDir",
                    help    = "input dir",
                    dest    = 'inputDir')
parser.add_argument("-i", "--inputFile",
                    help    = "input file",
                    dest    = 'inputFile',
                    default = '')
parser.add_argument("-f", "--fileFilter",
                    help    = "a list of comma separated string filters to apply across the input file space",
                    dest    = 'fileFilter',
                    default = '')
parser.add_argument("-l", "--fileFilterLogic",
                    help    = "the logic to apply across the file filter",
                    dest    = 'fileFilterLogic',
                    default = 'OR')
parser.add_argument("-d", "--dirFilter",
                    help    = "a list of comma separated string filters to apply across the input dir space",
                    dest    = 'dirFilter',
                    default = '')
parser.add_argument("-O", "--outputDir",
                    help    = "output image directory",
                    dest    = 'outputDir',
                    default = '')
parser.add_argument("--printElapsedTime",
                    help    = "print program run time",
                    dest    = 'printElapsedTime',
                    action  = 'store_true',
                    default = False)
parser.add_argument("--threads",
                    help    = "number of threads for innermost loop processing",
                    dest    = 'threads',
                    default = "0")
parser.add_argument("--outputLeafDir",
                    help    = "formatting spec for output leaf directory",
                    dest    = 'outputLeafDir',
                    default = "")
parser.add_argument("--test",
                    help    = "test",
                    dest    = 'test',
                    action  = 'store_true',
                    default = False)
parser.add_argument("-x", "--man",
                    help    = "man",
                    dest    = 'man',
                    action  = 'store_true',
                    default = False)
parser.add_argument("-y", "--synopsis",
                    help    = "short synopsis",
                    dest    = 'synopsis',
                    action  = 'store_true',
                    default = False)
parser.add_argument("--json",
                    help    = "output final return in json",
                    dest    = 'json',
                    action  = 'store_true',
                    default = False)
parser.add_argument("--overwrite",
                    help    = "overwrite files if already existing",
                    dest    = 'overwrite',
                    action  = 'store_true',
                    default = False)
parser.add_argument("--followLinks",
                    help    = "follow symbolic links",
                    dest    = 'followLinks',
                    action  = 'store_true',
                    default = False)
parser.add_argument("-v", "--verbosity",
                    help    = "verbosity level for app",
                    dest    = 'verbosity',
                    default = "1")
parser.add_argument('--version',
                    help    = 'if specified, print version number',
                    dest    = 'b_version',
                    action  = 'store_true',
                    default = False)

def main(argv=None):
    args = parser.parse_args()
    if args.man or args.synopsis:
        print(str_desc)
        if args.man:
            str_help     = synopsis(False)
        else:
            str_help     = synopsis(True)
        print(str_help)
        sys.exit(1)

    if args.b_version:
        print("Name:    %s\nVersion: %s" % (__pkg.name, __version__))
        sys.exit(1)

    args.str_version    = __version__
    args.str_desc       = synopsis(True)

    pf_do               = pfdo.pfdo(vars(args))

    # And now run it!
    # pudb.set_trace()
    d_pfdo              = pf_do.run(timerStart = True)

    if args.printElapsedTime:
        pf_do.dp.qprint(
                "Elapsed time = %f seconds" %
                d_pfdo['runTime']
        )

    return 0

if __name__ == "__main__":
    sys.exit(main())