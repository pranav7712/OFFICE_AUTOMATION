import os
import glob

ppt = os.getcwd()

filepath = "G:\OFFICE AUTOMATION FILES\Split Excel Files\Financial Sample.xlsx"

# breaking into different components of the Filepath
dir = os.path.dirname(filepath)  # this gives the directory name
extension = os.path.splitext(filepath)[1]  # provides the extension of the file
filename = os.path.splitext(filepath)[0]  # provides the name of the file in the Filepath



# Joining different components
joincheck = os.path.join(dir, filename+"_signed" +
                         extension)  # best part of join is that it automatically adds the / in the filepath



listcheck = os.listdir(dir)  # awesme because it simply provides a list of all the filenames with extension


# even you can join using the listdir command if you know the position of file through list indexing
joincheck_2 = os.path.join(dir, listcheck[0])


filesize = os.path.getsize(filepath)  # it provides the size of the file in Bytes


multiplefile = glob.glob(dir + "/*.xlsx")  # this will give list of multiple files through which we can iterate



print('The os.getcwd command gives this output ' + ppt)

print("This is the complete filepath : "+ filepath)

print("This is the directory name: " + dir)

print("This is the extension : "+ extension)

print("This is only the path till filename without extension: "+ filename)

print("The 3 inputs join to form complete filepath: "+ joincheck)

print(listcheck)

print('This is joining file through list indexing'+ joincheck_2)

print(filesize)

print(multiplefile)

for f in multiplefile:
    print("We can iterate like this..!!")
