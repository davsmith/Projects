"""Use dir and help functions to view information about objects"""
import math

print()

print('*** Print the docstring for the current module ***')
print(__doc__)
print()

print('*** Print the list of exports on the math module ***')
print(dir(math))
print()

print('*** Print the docstring for the cosine method ***')
print(help(math.cos))
print()

# Uncomment to see the math code in use.
# It has been commented out because it obscures the output from dir and help

# for degrees in range(0, 360):
#     print(
#         f"Degrees: {degrees},  cos: {math.cos(math.pi/180*degrees)},  sin: {math.sin(math.pi/180*degrees)}")
