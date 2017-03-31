#!/usr/bin/env python3

# open - opens the file readonly, unless you use 'w' for write, or 'a' for
# append
# close - closes the file
# read - reads the contents, so you can assign it to a variable
# readline - reads just one line
# truncate - empties the file - watch out!
# write('stuff') - writes 'stuff' to the file


from sys import argv

script, filename = argv

print("We're going to erase %s." % filename)
print("If you don't want that, hit CTRL-C (^C).")
print("If you do want that, hit ENTER")

temp = input('?')

print("Opening the file...")
target = open(filename, 'w')

print("Truncating the file. Goodbye!")
target.truncate()

print("Now I'm going to ask you for three lines.")

line1 = input("Line 1: ")
line2 = input("Line 2: ")
line3 = input("Line 3: ")

print("I'm going to write these to the file.")

target.write(line1)
target.write("\n")
target.write(line2)
target.write("\n")
target.write(line3)
target.write("\n")

print("And finally, we close it.")
target.close
