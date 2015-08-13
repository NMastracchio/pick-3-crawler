from pywinauto.application import Application
from msvcrt import getch
import shutil
import time
import os
  
#### Global variables ####

mode = 1 # 1=Fast, 2=Slow
dataFilesLoc = "C:/DATA FILES"
analysisLoc = "C:/ANALYSIS"
dataFileName = "Data3"
shortSleepTime = 0.1
longSleepTime = 1.5
# Rows 14-18
# Columns 18, 25, 30 35, 40, 47, 53, 58, 64, 69, 75
MDfilenames = [
  #"MD3.1",
  "MD3.2",
  "MD3.3",
  "MD3.4",
  "MD3.5",
  "MD3.6"
]
MDcolumns = [17, 24, 29, 34, 39, 46, 52, 57, 63, 68, 74]
# Rows 14-18
# Columns 18, 25, 29, 33, 38, 43, 48, 54, 58, 64, 69, 75
Wfilenames = [
  #"W3.1",
  "W3.2",
  "W3.3",
  "W3.4",
  "W3.5",
  "W3.6",
  #"W3.7"
]
Wcolumns = [17, 24, 28, 32, 37, 42, 47, 53, 57, 63, 68, 74]

#### Helper functions ####

#
# Gets the mode. Mode types are described in the call to raw_input
#
def getMode():
  while True:
    try:
      modeNum = int(raw_input("""
    Select a mode:
    [1]:  Fast mode - The program will run without pausing at each match found.
          The program will output a log file.

    [2]:  Slow mode - The program will pause at each match found.
          The program will still output a log file, but you must press [PAGE DOWN]
          for every match found to continue the program.

    Mode number: """),)
      if (modeNum != 1) and (modeNum != 2):
        print modeNum != 1
        raise ValueError("Please select either 1 or 2.")
    except ValueError:
      print "Please select either 1 or 2."
      continue
    else:
      break
  global mode
  mode = modeNum

#
# Gets the patterns that will be searched for. Returns a list of pattern lists
#
def getPatterns():
  patterns = []
  while True:
    try:
      numPatterns = int(raw_input("    Number of patterns: "))
    except ValueError:
      print "Please enter a valid number."
      continue
    else:
      break
  print """
    Enter the patterns that you would like to search for. For a blank, 
    enter an underscore. 

    (If more than 5 symbols are entered, only the first 5 will be used.)

    Example:
    + - _ + -"""
  for n in range(numPatterns):
    pattern = raw_input("\n    Pattern %s: " % str(n + 1))
    mappedPattern = map(str, pattern.split())
    mappedPattern = [char.replace('_', ' ') for char in mappedPattern]
    patterns.append(mappedPattern)
  print "Searching for the following patterns:"
  for n in range(numPatterns):
    if len(patterns[n]) > 5:
      patterns[n] = patterns[n][:5]
    print patterns[n]

  return patterns

#
# Returns the number of directories in dataFilesLoc
#
def getSubdirectories():
  return len([dir for dir in os.listdir(dataFilesLoc)
    if os.path.isdir(os.path.join(dataFilesLoc, dir))])

#
# Copies Data3 from dataFilesLoc to "C:/ANALYSIS" for file creation
#
def copyFile(dirNum):
  src = dataFilesLoc + "/" + str(dirNum) + "/" + dataFileName
  dest = analysisLoc + "/" + dataFileName
  print ("Copying %s/%s/%s to %s/%s: " % (dataFilesLoc, str(dirNum), 
                                          dataFileName, analysisLoc,
                                          dataFileName)),
  try:
    shutil.copy(src, dest)
  except shutil.Error as e:
    print('Error: %s' % e)
  except IOError as e:
    print('Error: %s' % e.strerror)
  print "Complete"

#
# Automatically runs Report3.exe and generates the output files
#
def runReport3():
  print "Running Report3.exe: ",
  app = Application.start("Report3.exe", create_new_console=True, 
                          wait_for_idle=False)
  window = app.top_window_()
  window.SetFocus()
  window.TypeKeys("s")
  for i in range(3):
    time.sleep(shortSleepTime)
    window.TypeKeys("{ENTER}")
  time.sleep(longSleepTime)
  for i in range(7):
    time.sleep(shortSleepTime)
    window.TypeKeys("{ENTER}")
  time.sleep(longSleepTime)
  for i in range(6):
    time.sleep(shortSleepTime)
    window.TypeKeys("{ENTER}")
  time.sleep(longSleepTime)
  window.TypeKeys("n") # Do not run again
  time.sleep(shortSleepTime)
  window.TypeKeys("x") # Exit program
  print "Complete"

#
# Searches for the patterns in each of the files in the filenames list.
# As per the specs, MD3.1, W3.1, & W3.7 are not included.
#
def extractPatterns(patterns):
  store = {}
  for filename in MDfilenames:
    store[filename] = {}
    fp = open(analysisLoc + "/" + filename)
    for i, line in enumerate(fp):
      if (i == 13) or (i == 14) or (i == 15) or (i == 16) or (i == 17):
        for n in range(len(MDcolumns)):
          charAppender(store, n, filename, line[MDcolumns[n]])
      elif i > 17:
        break
    fp.close()

  for filename in Wfilenames:
    store[filename] = {}
    fp = open(analysisLoc + "/" + filename)
    for i, line in enumerate(fp):
      if (i == 13) or (i == 14) or (i == 15) or (i == 16) or (i == 17):
        for n in range(len(Wcolumns)):
          charAppender(store, n, filename, line[Wcolumns[n]])
      elif i > 17:
        break
    fp.close()

  return store

#
# Replaces return characters, puts each character in the appropriate place
#
def charAppender(store, n, filename, character):
  if character == "\r":
    character = " "
  patternsArr = store[filename]["patterns"] = store[filename].get("patterns",[])
  try:
    patternsArr[n].append(character)
  except:
    patternsArr.append([])
    patternsArr[n].append(character)

#
# Takes the entered patterns and compares them to the extracted patterns
#
def findMatches(patterns, data, dirNum, file):
  def findMatchInList(toFind, inThis):
    for n in range(len(toFind)):
      if toFind[n] != inThis[n]:
        return False
      else:
        continue
    return True
  for pattern in range(len(patterns)): #for each pattern...
    for key, value in data.iteritems(): # for each key in data dict
      for n in range(len(data[key]["patterns"])):
        if findMatchInList(patterns[pattern], data[key]["patterns"][n]):
        #if patterns[pattern] == data[key]["patterns"][n]:
          if key[:2] == "MD":
          	colFormat = MDcolumns
          else:
          	colFormat = Wcolumns
          logText = """
    Match found!
    File: %s/%s
    Source file: %s/%s/%s
    Pattern found: %r
    Column #: %r
            """ % (analysisLoc, key, dataFilesLoc, dataFileName, str(dirNum), 
                    patterns[pattern], (colFormat[n] + 1))
          f.write(logText)
          print logText
          if mode == 2:
            while True:
              keyNum = ord(getch())
              if keyNum == 224: #Page Down
                break

getMode()
patterns = getPatterns()
numDirectories = getSubdirectories()
newpath = r'./logs/'
if not os.path.exists(newpath):
  os.makedirs(newpath)
logFilename = "./logs/log" + time.strftime("%m%d%y%H%M%S") + ".txt"
with open(logFilename,"a+") as f:
  for i in range(numDirectories):
    copyFile(i + 1)
    runReport3()
    data = extractPatterns(patterns)
    findMatches(patterns, data, (i+1), f )
print "Done! :D"