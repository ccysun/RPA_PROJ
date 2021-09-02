import psutil 

def find_process(name):
    for p in psutil.process_iter(["pid","name"]):
        if (p.info["name"] == name):
            return psutil.pid_exists(p.info["pid"])
    return False

def find_process2(name):
    return name in map(lambda p: p.name(), psutil.process_iter())

pp = find_process("iexplore.exe")
print (pp)

rst = find_process2("firefox.exe")
print(rst)



