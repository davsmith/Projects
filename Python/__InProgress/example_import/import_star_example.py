"""
Demonstrate how import * pollutes the namespace

Import from the main example file (i.e. do not run standalone)
"""

def first_function():
    """ First example function """
    print("This is the first function in the module to be imported")

def second_function():
    """ Second example function """
    print("This is the second function in the module to be imported")

def third_function():
    """ Third example function """
    print("This is the third function in the module to be imported")

if __name__ == "__main__":
    print('Not a standalone module')