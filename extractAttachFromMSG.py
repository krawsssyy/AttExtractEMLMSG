import os
import win32com.client
import hashlib

proc = []

def updateOutput(fn, zip_hash):
    with open('output.csv', 'a') as aa:
        aa.write(str(write_hash(fn)) + "," + zip_hash + '\n')
    return


def write_hash(FILENAME):
    with open(FILENAME, 'rb') as hh:
        m = hashlib.sha256()
        while True:
            data = hh.read(1)
            if not data:
                break
            m.update(data)
    return m.hexdigest()


def extract(filename):
    global proc
    output_count = 0
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.OpenSharedItem(filename)
        atts = msg.Attachments
        for att in atts:
            try:
                od = filename + "_output"
                os.path.exists(od) or os.makedirs(od)
                att.SaveAsFile(os.path.join(od, att.FileName))
                output_count += 1
                updateOutput(filename, write_hash(os.path.join(od, att.FileName)))
                proc.append(filename)
            except BaseException:
                print("Got error for one of the attachments for %s" % filename)
                continue
        if output_count == 0:
            print("No attachment found for file %s!" % f.name)     

    except IOError:
        print("Problem with %s or one of its attachments!" % f.name)
    return 1, output_count


if __name__=="__main__":
    path = 'C:\\Users\\heya\\Desktop\\test\\'
    files = [f for f in os.listdir(path) if ".msg" in f]
    for file in files:
        extract(path + file)
