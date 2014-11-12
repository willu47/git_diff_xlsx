from nose.tools import assert_equal, raises
from git_diff_xlsx import parse

@raises(ValueError)
def test_parse_xlsx():
    infile = ""
    outfile = ""
    actual = parse(infile, outfile)
    desired = ""
    assert_equal(actual, desired)
