""" Unit tests for localdb module.
"""

import unittest as ut
from unittest import mock

import localdb

def setUpModule():
    """ Sets up configuration for all test cases in this file.
    """
    pass


def tearDownModule():
    """ Tidies up configuration for all test cases in this file.
    """
    pass


class InstanceTestCase(ut.TestCase):

    @classmethod
    def setUpClass(cls):
        """ Sets up test configuration for all tests in this class.
        """
        pass

    @classmethod
    def tearDownClass(cls):
        """ Tidies up all test configurations for this class.
        """
        pass

    def setUp(self):
        """ Sets up each test method with default parameters.
        """
        self.mngr = localdb.InstanceManager()
        self.inst = self.mngr.create('TestInstance')

    def tearDown(self):
        """ Tidy up after each test method.
        """
        self.mngr.delete('TestInstance')

    def test_create(self):
        """ Creates (and deletes) a new LocalDB instance.
        """
        self.mngr.create('TempTestInstance')

        self.mngr.delete('TempTestInstance')

    def test_connstring(self):
        """ Correctly constructs instance database connection string.
        """
        cstr = self.inst.connection_string('mydatabase')
        expected = (
            r'Server={(LocalDB)\TestInstance};'
            r'Driver={ODBC Driver 13 for SQL Server};'
            r'Database={mydatabase};'
            r'Trusted_Connection=yes'
        )
        self.assertEqual(cstr, expected)
