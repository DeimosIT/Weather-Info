import time
import Foreca
import InfoRP5
import threading


dfs_foreca = []
list_city_name_foreca = []
dfs_rp5 = []
list_city_name_rp5 = []

def create_foreca():
    global dfs_foreca
    global list_city_name_foreca
    dfs_foreca, list_city_name_foreca = Foreca.main()

def create_rp5():
    global dfs_rp5
    global list_city_name_rp5
    dfs_rp5, list_city_name_rp5 = InfoRP5.main()

def main():
    global dfs_foreca
    global list_city_name_foreca
    global dfs_rp5
    global list_city_name_rp5
    t = time.time()
    th1 = threading.Thread(target=create_foreca)
    th2 = threading.Thread(target=create_rp5)
    th1.start()
    th2.start()
    th1.join()
    th2.join()
    Foreca.create_xlsx(dfs_foreca, list_city_name_foreca)
    InfoRP5.create_xlsx(dfs_rp5, list_city_name_rp5)
    print('time: ', time.time() - t)

if __name__ == '__main__':
    main()