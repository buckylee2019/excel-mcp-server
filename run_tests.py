#!/usr/bin/env python
"""
Run unit tests for Excel MCP Server pivot_xlwings module.
"""

import unittest
import sys

if __name__ == "__main__":
    # Run the pivot_xlwings tests
    test_suite = unittest.defaultTestLoader.loadTestsFromName('tests.test_pivot_xlwings')
    test_runner = unittest.TextTestRunner(verbosity=2)
    result = test_runner.run(test_suite)
    
    # Exit with non-zero code if tests failed
    sys.exit(not result.wasSuccessful())
