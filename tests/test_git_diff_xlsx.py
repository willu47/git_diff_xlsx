from nose.tools import assert_equal, raises
from ..git_diff_xlsx import parse
import sys


def test_parse_xlsx():
    infile = "tests/test1.xlsx"
    outfile = sys.stdout
    actual = parse(infile, outfile)
    desired = \
    """
    =================================
    Sheet: 1 [1, 1]
    =================================
    A1: Hello World
    """
    assert_equal(outfile, desired)
