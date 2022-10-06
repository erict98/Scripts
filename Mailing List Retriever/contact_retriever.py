import re
import os
import json

"""
Script creates a directory from a mailing list (from Outlook) and returns the list of email addresses
"""
directory = {}
path = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))

"""
Reads the JSON or writes to the JSON based on the mailing list.
"""
def create_dir(directory):
    mail = open(os.path.join(path, "mailing.txt"), "r")         # mailing list
    data = open(os.path.join(path, "directory.json"), "w+")     # directory

    # creates a JSON if empty or DNE
    if data.read(1): # reads the JSON and check if non-empty
        directory = json.load(data) 
        
    # writes the JSON
    else: 
        # splits the mailing list into <individual, email>
        for line in mail.readline().split("; "): 
            
            # splits the individual into <name> and <email>
            text = line.split("<") 
            name = text[0][0:-1]
            email = text[1][0:-1]
            
            # removes middle names
            name = name.replace(".", "")
            if re.search(" .", name[-2:]):
                name = name[:-2]
                
            # removes (nicknames)
            if "(" in name:
                posA = name.index("(")
                posB = name.index(")")
                name = name[0:posA-1] + name[posB+1:]
            
            # updates the directory
            directory[name] = email

        data.write(json.dumps(directory)) 

""" 
Reads the input of names and returns output of email addresses.
Prints "ERROR: - <name>" if cannot be found.
"""
def read_dir(directory):
    input = open(os.path.join(path, "input.txt"), "r")
    output = open(os.path.join(path, "output.txt"), "w")
    output.truncate()
    for text in input:
        
        # removes newlines and middle names
        name = text.strip()
        if re.search(" .", name[-2:]):
            name = name[:-2]

        if name in directory:
            output.write(directory[name] + "\n")
        else:
            # checks for abbreviated names (e.g., Edward -> Ed, not applicable for William -> Billy)
            for n in range(len(name)):
                if name[:-1*n] in directory:
                    output.write(directory[name[:-1*n]] + "\n")
                    break
                elif n == len(name) -1:
                    output.write("ERROR: " + name + "\n")
                
if __name__ == "__main__":
    create_dir(directory)
    read_dir(directory)
