str_description = """
    This module provides some very simple shell-based job running
    methods.
"""


import  subprocess
import  os
import  pudb
import  json

class jobber:

    def __init__(self, d_args : dict):
        """Constructor for the jobber class.

        Args:
            d_args (dict): a dictionary of "arguments" (parameters) for the
                           object.
        """
        self.args   = d_args.copy()
        if not 'verbosity'      in self.args.keys(): self.args['verbosity']     = 0
        if not 'noJobLogging'   in self.args.keys(): self.args['noJobLogging']  = False

    def dict2JSONcli(self, d_dict : dict) -> str:
        """Convert a dictionary into a CLI conformant JSON string.

        An input dictionary of

            {
                'key1': 'value1',
                'key2': 'value2'
            }

        is converted to a string:

            "{\"key1\":\"value1\",\"key2\":\"value2\"}"

        Args:
            d_dict (dict): a python dictionary to convert

        Returns:
            str: CLI equivalent string.
        """

        str_JSON    = json.dumps(d_dict)
        str_JSON    = str_JSON.replace('"', r'\"')
        return str_JSON

    def dict2cli(self, d_dict : dict) -> str:
        """Convert a dictionary into a CLI conformant JSON string.

        An input dictionary of

            {
                'key1': 'value1',
                'key2': 'value2',
                'key3': true,
                'key4': false
            }

        is converted to a string:

            "--key1 value1 --key2 value2 --key3"

        Args:
            d_dict (dict): a python dictionary to convert

        Returns:
            str: CLI equivalent string.
        """
        str_cli     : str = ""
        for k,v in d_dict.items():
            if type(v) == bool:
                if v:
                    str_cli += '--%s ' % k
            elif len(v):
                str_cli += '--%s %s ' % (k, v)
        return str_cli

    def job_run(self, str_cmd: str):
        """
        Running some CLI process via python is cumbersome. The typical/easy
        path of

                            os.system(str_cmd)

        is deprecated and prone to hidden complexity. The preferred
        method is via subprocess, which has a cumbersome processing
        syntax. Still, this method runs the `str_cmd` and returns the
        stderr and stdout strings as well as a returncode.
        Providing readtime output of both stdout and stderr seems
        problematic. The approach here is to provide realtime
        output on stdout and only provide stderr on process completion.
        """
        d_ret       : dict = {
            'stdout':       "",
            'stderr':       "",
            'cmd':          "",
            'cwd':          "",
            'returncode':   0
        }
        str_stdoutLine  : str   = ""
        str_stdout      : str   = ""

        p = subprocess.Popen(
                    str_cmd.split(),
                    stdout      = subprocess.PIPE,
                    stderr      = subprocess.PIPE,
        )

        # Realtime output on stdout
        while True:
            stdout      = p.stdout.readline()
            if p.poll() is not None:
                break
            if stdout:
                str_stdoutLine = stdout.decode()
                if int(self.args['verbosity']):
                    print(str_stdoutLine, end = '')
                str_stdout      += str_stdoutLine
        d_ret['cmd']        = str_cmd
        d_ret['cwd']        = os.getcwd()
        d_ret['stdout']     = str_stdout
        d_ret['stderr']     = p.stderr.read().decode()
        d_ret['returncode'] = p.returncode
        if int(self.args['verbosity']) and len(d_ret['stderr']):
            print('\nstderr: \n%s' % d_ret['stderr'])
        return d_ret

    def job_runbg(self, str_cmd : str) -> dict:
        """Run a job in the background

        Args:
            str_cmd (str): CLI string to run

        Returns:
            dict: a dictionary of exec state
        """
        d_ret       : dict = {
            'uid'       : "",
            'cmd'       : "",
            'cwd'       : ""
        }

        process = subprocess.Popen(
                    str_cmd.split(),
                    stdout      = subprocess.PIPE,
                    stderr      = subprocess.PIPE,
                    close_fds   = True
        )

        for line in process.stdout:
            pass
        process.wait()

        d_ret['uid']        = str(os.getuid())
        d_ret['cmd']        = str_cmd
        d_ret['cwd']        = os.getcwd()
        return d_ret

    def job_stdwrite(self, d_job : dict, str_outputDir : str, str_prefix : str = "") -> dict:
        """
        Capture the d_job entries to respective files.
        """
        if not self.args['noJobLogging']:
            for key in d_job.keys():
                with open(
                    '%s/%s%s' % (str_outputDir, str_prefix, key), "w"
                ) as f:
                    f.write(str(d_job[key]))
                    f.close()
        return {
            'status': True
        }
