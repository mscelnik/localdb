""" SQL Server LocalDB instance management for Python.

WINDOWS ONLY!

This module provides a wrapper for the Windows command line application
SQLLocalDB.exe.  This allows you to manage LocalDB instances on your computer.

Classes:
    CmdExecutor: Interface to the Windows sqllocaldb.exe command line tool.  Do
        not use this directly; instead use InstanceManager to create and manage
        instances.
    Instance: SQL Server LocalDB instance.  Do not instantiate directly; instead
        use the InstanceManager to create Instance objects.
    InstanceManager: Manages LocalDB instances.
"""

from collections import namedtuple

SQL_ATTACH = """
CREATE DATABASE [{dbname}]
ON (FILENAME=N'{fpath}')
FOR ATTACH_REBUILD_LOG
"""

SQL_DETACH = """
ALTER DATABASE [{dbname}] SET OFFLINE;
USE master EXEC sp_detach_db @dbname = N'{dbname}';
"""

ExecutableInfo = namedtuple(
    'ExecutableInfo',
    'path version regkey',
)


InstanceInfo = namedtuple(
    'InstanceInfo',
    'name version shared_name owner auto_create state last_start pipe_name',
)


ServerVersion = namedtuple(
    'ServerVersion',
    'name version major minor micro build'
)


class CmdExecutor(object):
    """ Interface to the SQLLocalDB.exe command line application on Windows.

    WARNINGS
    --------
    1.  WINDOWS ONLY!
    2.  DO NOT USE.  The Instance and InstanceManager classes use this
            class.  You should not call it directly - use those other classes
            instead.
    """

    def __init__(self):
        """ Initialize the executor by finding the local he SQLLocalDB.exe.
        """
        # We don't need to store all the available SQLLocalDB.exe paths;
        # here we just use the most recent.
        exes = self.find_exes()
        if exes is not None:
            self._info = exes[0]
        else:
            self._info = None

    @property
    def available(self):
        """ True if SQLLocalDB is installed on this computer, otherwise False.
        """
        return self._info is not None

    @property
    def exe_path(self):
        """ Path to the SQLLocalDB.exe file on this computer, if installed.
        """
        if self.available:
            return self._info.path
        else:
            return None

    def find_exes(self):
        """ Locates available SQLLocalDB.exe files on the host computer.

        Currently this function does not search.  It assumes the file exists
        and is on the PATH.

        TODO: Use Windows registry to locate installed SQLLocalDB.exe files.

        Returns:
            List of ExecutableInfo objects, sorted by version descending.
            Empty list if LocalDB is not installed.
        """
        fake = ExecutableInfo(path='SQLLocalDB.exe', version='X.X', regkey='')
        return [fake]

    def call(self, cmd, **kwargs):
        """ Calls the SQLLocalDB.exe program and returns the output string.

        Args:
            cmd (str): SQLLocalDB operation command.  One of: create|c,
                delete|d, start|s, stop|p, share|h, unshare|u, info|i,
                versions|v, trace|t.
            kwargs: Arguments for command.  Exact arguments depend on the
                command as presented in the table below.

        |-----------|----------------|-----------------------------------------|
        | Command   | Kwargs         | Notes                                   |
        |-----------|----------------|-----------------------------------------|
        | create/c  | name           | String.  Instance name.                 |
        |           | version        | String, optional. Defaults to most      |
        |           |                |   recent installed version if omitted.  |
        |           | start          | Bool, optional. Does NOT start if       |
        |           |                |   ommited.                              |
        |           |                |                                         |
        |           |                | RETURNS: None                           |
        |-----------|----------------|-----------------------------------------|
        | delete/d  | name           | String.  Instance name.                 |
        |           |                |                                         |
        |           |                | RETURNS: None                           |
        |-----------|----------------|-----------------------------------------|
        | start/s   | name           | String.  Instance name.                 |
        |           |                |                                         |
        |           |                | RETURNS: None                           |
        |-----------|----------------|-----------------------------------------|
        | stop/p    | name           | String.  Instance name.                 |
        |           | nowait         | Bool, optional. Set True to request     |
        |           |                |   shutdown with NOWAIT option.          |
        |           | kill           | Bool, optional. Set True to kill the    |
        |           |                |   process without contacting it.        |
        |           |                |                                         |
        |           |                | RETURNS: None                           |
        |-----------|----------------|-----------------------------------------|
        | share/h   | name           | String.  Instance name                  |
        |           | sharedname     | String. Name of share.                  |
        |           | owner          | String, optional. User ID owning share. |
        |           |                |                                         |
        |           |                | RETURNS: None                           |
        |-----------|----------------|-----------------------------------------|
        | unshare/u | sharedname     | String, share name to stop sharing.     |
        |           |                |                                         |
        |           |                | RETURNS: None                           |
        |-----------|----------------|-----------------------------------------|
        | info/i    | name           | String.  Instance name.                 |
        |           |                |                                         |
        |           |                | RETURNS: List of instance names, if no  |
        |           |                |   argument. Instance Info tuple if      |
        |           |                |   valid instance name provided.         |
        |-----------|----------------|-----------------------------------------|
        | versions/ | N/A            | RETURNS: List of ServerVersion tuples.  |
        | v         |                |                                         |
        |           |                |                                         |
        |-----------|----------------|-----------------------------------------|
        | trace/t   | enable         | Bool.  Set True to turn on, False to    |
        |           |                |   turn off.                             |
        |           |                |                                         |
        |           |                | RETURNS: None                           |
        |-----------|----------------|-----------------------------------------|
        """
        import sys
        import subprocess

        if not sys.platform.startswith('win'):
            raise RuntimeError('This function only works on Windows!')

        if not self.available:
            raise RuntimeError('SQLLocalDB.exe is not installed.')

        EXE = self.exe_path

        # Don't want case sensitivity issues.
        cmd = cmd.lower()

        if cmd == 'create':
            name = kwargs.get('name', None)
            if name is None:
                raise ValueError(
                    'Must give instance name for "create" operation.')
            version = kwargs.get('version', '')
            if kwargs.get('start', False):
                start = '-s'
            else:
                start = ''
            stmt = f'{EXE} create "{name}" {version} {start}'

        elif cmd in ('delete', 'start'):
            name = kwargs.get('name', None)
            if name is None:
                raise ValueError(
                    f'Must give instance name for "{cmd}" operation.')
            stmt = f'{EXE} {cmd} "{name}"'

        elif cmd == 'stop':
            name = kwargs.get('name', None)
            if name is None:
                raise ValueError(
                    f'Must give instance name for "stop" operation.')
            nowait = kwargs.get('nowait', None)
            kill = kwargs.get('kill', None)
            stmt = f'{EXE} {cmd} "{name}"'

        elif cmd == 'share':
            name = kwargs.get('name', None)
            if name is None:
                raise ValueError(
                    f'Must give instance name for "share" operation.')
            sharedname = kwargs.get('sharedname', None)
            owner = kwargs.get('owner', None)
            stmt = f'{EXE} {owner} "{name}" "{sharedname}"'

        elif cmd == 'unshare':
            sharedname = kwargs.get('sharedname', None)
            if sharedname is None:
                raise ValueError(
                    f'Must give share name for "unshare" operation.')
            stmt = f'{EXE} unshare "{sharedname}"'

        elif cmd == 'info':
            name = kwargs.get('name', '')
            stmt = f'{EXE} info {name}'

        elif cmd == 'versions':
            stmt = f'{EXE} versions'

        elif cmd == 'trace':
            enable = kwargs.get('enable', None)
            if enable is None:
                raise ValueError(
                    f'Must give enable flag for "trace" operation.')
            if enable:
                enable = 'on'
            else:
                enable = 'off'
            stmt = f'{EXE} trace {enable}'

        else:
            raise NotImplementedError(
                'SQLLocalDB operation "{cmd}" not yet implemented.')

        proc = subprocess.Popen(
            stmt,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            universal_newlines=True,
        )
        proc.wait()
        data = proc.communicate()
        return data[0]


class Instance(object):
    """ Represents a LocalDB instance on the host computer.

    Warning!
    --------
    Do not create instances of this class directly.  Use the InstanceManager
    to create instead.
    """

    def __init__(self, info):
        self._info = info
        self._exe = CmdExecutor()

    @property
    def name(self):
        return self._info.name

    @property
    def version(self):
        return self._info.version

    def start(self):
        self._exe.call('start', name=self.name)
        self._info = self._exe.call('info', name=self.name)

    def stop(self):
        self._exe.call('stop', name=self.name)
        self._info = self._exe.call('info', name=self.name)

    def share(self, sharedname, owner=None):
        self._exe.call('share', name=self.name, sharedname=sharedname,
                       owner=owner)

    def unshare(self, sharedname):
        self._exe.call('unshare', sharedname=sharedname)

    def info(self):
        return self._info

    def reset(self):
        """ Stops, deletes and recreates the instance. Use to detach all DBs.
        """
        self._exe.call('stop', name=self.name)
        self._exe.call('delete', name=self.name)
        self._exe.call(
            'create',
            name=self.name,
            version=self.version,
            start=True,
        )

    def _is64bit(self):
        """ Determines if Python is running in 64-bit mode.
        """
        import sys
        return sys.maxsize > 2**32

    def _all_drivers(self):
        """ Searches the Windows Registry for all installed ODBC drivers.

        If running 64-bit Python, only looks for 64-bit drivers, and same for
        32-bit Python.

        Returns:
            List of driver names.
        """
        import winreg

        drivers = []

        # Connect to the relevant registry key.
        REG_KEY = r'SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers'
        if self._is64bit():
            accessmode = winreg.KEY_WOW64_64KEY + winreg.KEY_QUERY_VALUE
        else:
            accessmode = winreg.KEY_WOW64_32KEY + winreg.KEY_QUERY_VALUE
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_KEY, accessmode)
        except FileNotFoundError:
            return None

        # Get all installed ODBC drivers from the registry.
        i = 0
        keeplooking = True
        while keeplooking:
            try:
                name, data, typ = winreg.EnumValue(key, i)
                drivers.append(name)
                i += 1
            except OSError:
                # OSError occurs when there are no more values to read.
                keeplooking = False

        winreg.CloseKey(key)

        return drivers

    def _valid_drivers(self, drivers):
        """ Filters a list of driver names to those compatible with LocalDB.

        The valid driver names are sorted, newest to oldest.

        Args:
            drivers (list of str): List of driver names, e.g. as generated from
                the _find_drivers method.

        Returns:
            List of valid driver names for LocalDB.
        """
        import re
        valid_drivers = []
        patterns = [
            r'ODBC Driver [0-9]{2} for SQL Server',  # Newer drivers, preferred.
            r'SQL Server Native Client 11\.0',  # Older, SQL Server 2012 only.
        ]
        alldrivers = '\n'.join(drivers)
        for pattern in patterns:
            matches = re.findall(pattern, alldrivers)
            valid_drivers.extend(sorted(matches, reverse=True))
        return valid_drivers

    def latest_driver(self):
        """ Finds the latest installed LocalDB compatible driver on the host.

        This returns the driver name in a form suitable for ODBC connection
        strings, e.g. for use with pyodbc.

        At time of writing, LocalDB compatible drivers are:
            ODBC Driver 17 for SQL Server = shipped with SQL Server 2017 (maybe)
            ODBC Driver 13 for SQL Server = shipped with SQL Server 2016
            ODBC Driver 11 for SQL Server = shipped with SQL Server 2014
            SQL Server Native Client 11.0 = shipped with SQL Server 2012

        All drivers are backwards compatible so we only need the latest driver,
        regardless of the LocalDB instance version.

        These earlier drivers do NOT support LocalDB:
            SQL Server Native Client 10.0 = shipped with SQL Server 2008
            SQL Server Native Client 9.0  = shipped with SQL Server 2005

        See these URLs for more info:
            - https://docs.microsoft.com/en-us/sql/connect/connect-history
            - https://docs.microsoft.com/en-us/sql/connect/odbc/windows/
                release-notes
            - https://docs.microsoft.com/en-us/sql/relational-databases/
                native-client/sql-server-native-client
        """
        drivers = self._all_drivers()
        valid_drivers = self._valid_drivers(drivers)
        return valid_drivers[0]

    def connection_string(self, dbname=None):
        """ Returns a valid DSN connection string for an instance database.

        This method does NOT check is the database is currently attached to this
        LocalDB instance.  If it is not, then you will not be able to connect
        to it.

        Args:
            dbname (str): Optional, valid database name in the instance.  If
                not given, then the connection string is effectively for the
                instance's master' database.

        Return:
            DSN connection string, e.g. for use in pyodbc.
        """
        server = f'Server={{(LocalDB)\\{self.name}}}'
        driver = f'Driver={{{self.latest_driver()}}}'
        if dbname is not None:
            database = f'Database={{{dbname}}}'
        else:
            database = ''
        sec = 'Trusted_Connection=yes'
        user = ''
        parts = filter(
            lambda s: len(s) > 0,
            [server, driver, database, user, sec]
        )
        return ';'.join(parts)

    def url(self, dbname):
        """ Returns a sqlalchemy engine URL for an instance database.

        Args:
            dbname (str): Valid database name in the instance.

        Returns:
            sqlalchemy connection URL.

        """
        import urllib.parse
        dsn = self.connection_string(dbname)
        dsn = urllib.parse.quote_plus(dsn)
        return f'mssql+pyodbc:///?odbc_connect={dsn}&autocommit=true'

    # I haven't decided is database attachment/detachment should be part of this
    # interface.  It requires a ODBC driver package like pyodbc to work.

    def attach(self, filepath, dbname=None):
        """ Attaches a MDF file to a database within the instance.

        Uses Transact-SQL to attached the database.  Any SQL errors get raised
        as LocalDBError() exceptions.

        Args:
            filepath (str): Full path to MDF file.
            dbname (str): [Optional] name to give database.  If omitted, then
                uses the filename without extension.

        Returns:
            The database name inside the instance if successful.
        """
        try:
            import pyodbc
        except ImportError:
            raise

        if dbname is None:
            dbname, _ = os.path.splitext(os.path.basename(filepath))

        try:
            dsn = self.connection_string()
            conn = pyodbc.connect(dsn, autocommit=True)
            sql = SQL_ATTACH.format(dbname=dbname, fpath=filepath)
            conn.execute(sql)
            conn.close()
        except pyodbc.ProgrammingError as e:
            info = parse_error(e.args[1])
            raise LocalDBError(
                info['SHORT'],
                description=info['MSG'],
                solution=info['SOLUTION']) from e
        except pyodbc.Error as e:
            info = parse_error(e.args[1])
            raise LocalDBError(
                info['SHORT'],
                description=info['MSG'],
                solution=info['SOLUTION']) from e

        return dbname

    def detach(self, dbname):
        try:
            import pyodbc
        except ImportError:
            raise

        try:
            dsn = self.connection_string()
            conn = pyodbc.connect(dsn, autocommit=True)
            sql = SQL_DETACH.format(dbname=dbname)
            conn.execute(sql)
            conn.close()

        except RuntimeError as e:
            raise LocalDBError('Failed to detach SQL database!') from e


class InstanceManager(object):
    """ Manages installed LocalDB instances on the host computer.
    """

    def __init__(self):
        self.exe = CmdExecutor()
        self._instances = self._findall()

    def _findall(self):
        """ Gets a reference to all installed instances on this computer.

        Returns:
            Dictionary of Instance objects.  The keys are the lower-case
            instance names.
        """
        instances = {}
        names = self.info()
        for name in names:
            info = self.info(name)
            inst = Instance(info)
            instances[name.lower()] = inst
        return instances

    def get(self, name, create=False):
        """ Returns Instance given valid name.  Optionally creates instance.

        Args:
            name (str): Valid LocalDB instance name.
            create (bool): Optional.  Set to True to automatically create the
                instance if it does not exist.

        Returns:
            Instance object if name is valid, otherwise None
        """
        lowername = name.lower()
        inst = self._instances.get(lowername, None)
        if inst is None:
            # Search for instances on this computer, if we have no record of
            # the requested instance.
            names = [n.lower() for n in self.info()]
            if lowername in names:
                info = self.info(name)
                inst = Instance(info)
                self._instances[name] = inst
            elif create:
                return self.create(name)
        return inst

    def create(self, name, version='', start=False):
        """ Create a new LocalDB instance with specified name and version.

        Args:
            name (str): Valid LocalDB instance name.
            version (str): Version number - must specify at least the major AND
                minor versions, e.g. '12.0'.  If omitted defaults to the latest
                installed version.
            start (bool): Set to True to automatically start the new instance.
                If omitted, the instance is NOT started.

        Returns:
            Reference to an Instance object if successful.
        """
        self.exe.call('create', name=name, version=version, start=start)

        # Save a reference to this instance for later use, and return it.
        info = self.info(name)
        inst = Instance(info)
        self._instances[name.lower()] = inst
        return inst

    def delete(self, name):
        """ Stops and deletes the named LocalDB instance, if it exists.

        Args:
            name (str): Valid LocalDB instance name.
        """
        self.stop(name)
        self.exe.call('delete', name=name)
        self._instances.pop(name.lower())

    def start(self, name):
        """ Starts the named LocalDB instance, if it exists.

        Args:
            name (str): Valid LocalDB instance name.
        """
        self.exe.call('start', name=name)

    def stop(self, name):
        """ Stops the named LocalDB instance, if it exists.

        Args:
            name (str): Valid LocalDB instance name.
        """
        self.exe.call('stop', name=name)

    def share(self, name, sharedname, owner=None):
        """ Shares the named LocalDB instance to the share name.

        Args:
            name (str): Valid LocalDB instance name.
            sharedname (str): Name for the share.
            owner (str): Optional, user or account UID.
        """
        self.exe.call('share', name=name, sharedname=sharedname, owner=owner)

    def unshare(self, sharedname):
        """ Unshares the shared LocalDB instance given the share name.

        Args:
            sharedname (str): Valid LocalDB instance shared name.
        """
        self.exe.call('unshare', sharedname=sharedname)

    def info(self, name=''):
        """ Returns information about LocalDB instances on the local computer.

        If you supply an instance name, this method returns an InstanceInfo
        for that instance.  Otherwise, it returns a list of all installed
        instance names.

        Args:
            name (str): Optional.  If supplied, this method returns information
                about that specific instance.

        Returns:
            (where name=='') List of instance names.
            (where name==instance name) InstanceInfo for LocalDB instance.
        """
        data = self.exe.call('info', name=name)
        lines = data.splitlines()
        if name is None or name == '':
            return lines
        else:
            keys = {
                'name': 'name',
                'version': 'version',
                'shared name': 'shared_name',
                'owner': 'owner',
                'auto-create': 'auto_create',
                'state': 'state',
                'last start time': 'last_start',
                'instance pipe name': 'pipe_name',
            }
            params = {}
            for line in lines:
                if (len(line) == 0) or (':' not in line):
                    continue
                key, value = line.split(':', maxsplit=1)
                mappedkey = keys[key.lower().strip()]
                params[mappedkey] = value.strip()
            return InstanceInfo(**params)

    def versions(self):
        """ Returns list of LocalDB versions installed on the host computer.

        Each list item is a ServerVersion namedtuple instance.
        """
        import re
        pattern = (
            r'(?P<name>.*?)\((?P<version>(?P<major>[0-9]+?)\.'
            r'(?P<minor>[0-9]+?)\.(?P<micro>[0-9]+?)\.(?P<build>[0-9]+?))\)'
        )
        data = self.exe.call('versions')
        lines = data.splitlines()
        vs = []
        for line in lines:
            matches = re.search(pattern, line)
            if matches is not None:
                v = ServerVersion(**matches.groupdict())
                vs.append(v)
        return vs

    def trace(self, enable=True):
        self.exe.call('trace', enable=enable)


class LocalDBError(RuntimeError):

    def __init__(self, msg, *args, **kwargs):
        """ Initialize the LocalDB error.

        Args:
            msg = Short description.

        Kwargs:
            solution = Potential solutions, if any.
            description = Long description.
        """
        super().__init__(self, msg, *args)
        self.solution = kwargs.get('solution', '')
        self.description = kwargs.get('description', '')
        self.short_description = msg

    def __repr__(self):
        return self.summary

    @property
    def summary(self):
        """ Returns a string summarizing all exception information.
        """
        return (
            'Error details\n--------------------\n'
            '{desc}\n'
            'Possible solutions\n------------------\n{soln}\n\n'
        ).format(desc=self.description, soln=self.solution)


def parse_error(msg):
    """ Parses MS SQL Server error messages into a dictionary.

    The message information is returned as a dictionary with keys:
        - SQLSTATE
        - DRIVER
        - DESCRIPTION = The detailed error description.
        - CODE = The SQL Server error code
        - CMD
        - SOLUTION = Possible solutions to prevent the error.
        - MSG = All error information in a single, multi-line string.
        - SHORT = Short description (one-liner)

    See URL below for SQL Server error codes:
        https://technet.microsoft.com/en-us/library/cc645603(v=sql.105).aspx

    Args:
        msg = String error message returned by DB-API (e.g. pyodbc).
    """
    import re
    regex = (
        r'\[(?P<SQLSTATE>.*?)\] (?P<DRIVER>\[.*\])(?P<DESC>.*?) '
        r'\((?P<CODE>[0-9]+?)\) \((?P<CMD>.+?)\)'
    )
    m = re.match(regex, msg)
    if m:
        info = {
            'SQLSTATE': m.group('SQLSTATE'),
            'DRIVER': m.group('DRIVER').strip(),
            'DESCRIPTION': m.group('DESC').strip(),
            'CODE': int(m.group('CODE')),
            'CMD': m.group('CMD'),
            'SOLUTION': None,
        }
        info['SHORT'] = (
            'SQL Server failed with error {}.'.format(info['CODE'])
        )

        # Add potential solutions to the info dict, if we have any.
        if info['CODE'] in (5120, 5121, 5123, 5133):
            # Access denied!
            info['SOLUTION'] = (
                'This SQL Server MDF file is NOT attachable. '
                'Possible solutions:\n'
                '  1. Ensure you have full read/write access to '
                'the MDF file AND its parent folder.\n'
                '  2. Check file is not already attached '
                '(using SQL Server Management Studio.)\n'
                '  3. Delete the corresponding LDF file.\n'
                '  4. Copy MDF file to a new location.\n')

        elif info['CODE'] == 823:
            info['SOLUTION'] = (
                'The MDF file is corrupt or not a genuine database file. Try '
                'to re-create the current file or use a different file.')

        elif info['CODE'] == 5118:
            info['SOLUTION'] = (
                'SQL Server cannot read MDF files in compressed folders. '
                'Ensure the target folder is not compressed (text appears '
                'blue in Windows Explorer.)  Edit the folder properties to '
                'uncompress it and its contents.')

        elif "Server does not exist or access denied" in info['DESCRIPTION']:
            info['SOLUTION'] = (
                'Could not access SQL Server on your computer.\n'
                'Potential solutions:\n'
                '  1. Install SQL Server Express With Tools (2014+) from:\n'
                '     https://www.microsoft.com/en-us/server-cloud/products/'
                'sql-server-editions/sql-server-express.aspx.\n'
                '  2. Contact Qraken team for further assistance.\n\n'
            )
        else:
            info['SOLUTION'] = (
                'No known solutions available.'
            )

        info['MSG'] = (
            # '\nError details\n-------------\n'
            '  SQL Driver:     {DRIVER}\n'
            '  SQL state code: {SQLSTATE}\n'
            '  Error code:     {CODE}\n\n'
            'Long description\n----------------\n{DESCRIPTION}\n'
            # 'Possible solutions:\n-------------------\n{SOLUTION}\n'
        ).format(**info)

        return info
    else:
        return None


if __name__ == '__main__':
    mngr = InstanceManager()
    inst = mngr.get('Qraken')
    print(inst.connection_string('mydatabase'))
