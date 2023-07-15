import glob
import os
import email
from email import policy
from multiprocessing import Pool
import zipfile
import subprocess
import hashlib


def updateOutput(fn, zip_hash):
    with open('output_zip.csv', 'a') as aa:
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
    if filename == "extractZIPsFromEMLAndExtractAtt.py" or filename == "output_zip.csv":
        print("Reached script/output file!Continuing...")
        return 1, 0
    output_count = 0
    try:
        with open(filename, "rb") as f:
            msg = email.message_from_bytes(f.read(), policy=policy.default)
            od = filename + "_output"
            os.path.exists(od) or os.makedirs(od)
            for attachment in msg.iter_attachments():
                try:
                    output_filename = attachment.get_filename()
                except AttributeError:
                    print("Got string instead of filename for %s. Skipping." % f.name)
                    continue
                if output_filename:
                    with open(os.path.join(od, output_filename), "wb") as of:
                        try:
                            of.write(attachment.get_payload(decode=True))
                            output_count += 1
                        except TypeError:
                            print("Couldn't get payload for %s" % output_filename)
                    if zipfile.is_zipfile(os.path.join(od, output_filename)):
                        zip_hash = write_hash(os.path.join(od, output_filename))
                        updateOutput(filename, zip_hash)
                        subprocess.call(["C:\\Program Files\\7-Zip\\7z.exe", "x", os.path.join(od, output_filename) ,"-o"+"C:\\Users\\heya\\Desktop\\idk1\\"+zip_hash+"\\"]) # call to extract to location specified by -o
            if output_count == 0:
                print("No attachment found for file %s!" % f.name)     
        if output_filename and zipfile.is_zipfile(os.path.join(od, output_filename)):
            subprocess.call(['rm', filename]) # delete file for which we succeeded in order to keep track of what's done

    except IOError:
        print("Problem with %s or one of its attachments!" % f.name)
    return 1, output_count

if __name__=="__main__":
    pool = Pool(None)
    res = pool.map(extract, glob.iglob("*"))
    pool.close()
    pool.join()
