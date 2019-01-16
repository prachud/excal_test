import os
import zipfile
from zipfile import ZipFile

def zipdir(path,file,groupID):
    fileName = "{}.zip".format(groupID)
    os.chdir(path)
    zipf = ZipFile(fileName, 'w',zipfile.ZIP_DEFLATED)
    zipf.write(file)
    zipf.close()

"""                
if __name__ == '__main__':
    zipf = zipfile.ZipFile('test.zip', 'w', zipfile.ZIP_DEFLATED)
    zipdir('tmp/', zipf)
    zipf.close()"""