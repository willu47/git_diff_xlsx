from setuptools import setup, find_packages
from setuptools.command.test import test as TestCommand


class NoseTestCommand(TestCommand):

    def finalize_options(self):
        TestCommand.finalize_options(self)
        self.test_args = []
        self.test_suite = True

    def run_tests(self):
        # Run nose ensuring that argv simulates running nosetests directly
        import nose
        nose.run_exit(argv=['nosetests'])

setup(

	name = "git_diff_xlsx",
	version = "0.1",
	packages = find_packages(exclude=['*test']),
	scripts = [],
	tests_require=['nose'],
	dependency_links=
	   ["https://raw.githubusercontent.com/willu47/pycel/master/src/pycel/tokenizer.py",
	   "https://raw.githubusercontent.com/willu47/pycel/master/src/pycel/excelutil.py"],

        cmdclass={'test': NoseTestCommand},

	install_requires = ['lxml'],

	author = "Will Usher",
	author_email = "w.usher@ucl.ac.uk",
	description = "Converts Microsoft Excel files to text to enable easy \
	               git diff",
	license=open('license.txt').read(),
	keywords = "git excel",
	url = "https://github.com/willu47/git_diff_xlsx",

)
