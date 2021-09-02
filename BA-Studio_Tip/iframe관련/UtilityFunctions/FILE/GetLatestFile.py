import os

def getLatestFile(folder_name):

   files = [os.path.join(folder_name, x) for x in os.listdir(folder_name) if x.endswith(".xlsx")]
   #files = [os.path.join(folder_name, x) for x in os.listdir(folder_name)]

   latest = max(files , key = os.path.getctime)

   return(latest)


#Driver code
latestFile = getLatestFile('C:\\temp')
print(latestFile)
