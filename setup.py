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
	license=open('license.txt').read(),
	scripts = [],
	tests_require=['nose'],

        cmdclass={'test': NoseTestCommand},

	install_requires = ['xlrd'],

	author = "Will Usher",
	author_email = "w.usher@ucl.ac.uk",
	description = "Converts Microsoft Excel files to text to enable easy \
	               git diff",
	license = "",
	keywords = "git excel",
	url = "https://github.com/willu47/git_diff_xlsx",

)
