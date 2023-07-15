import glob
from multiprocessing import Pool

def check(filename):
    if filename == "checkIfZipIsPassProt.py" or filename == "output.txt":
        print("Reached script/output file!Continuing...")
        return 
    try:
        with open(filename, "rb") as f, open ('output.txt', 'a') as g:
            data = f.read(6)
            flags = f.read(2)
            if int.from_bytes(flags, 'little') % 2 == 1:
                g.write("%s\n" % filename)
            else:
                print("File %s isn't password protected!\n" % filename)
    except e:
        print(e)
            

if __name__=="__main__":
    pool = Pool(None)
    res = pool.map(check, glob.iglob("*"))
    pool.close()
    pool.join()


