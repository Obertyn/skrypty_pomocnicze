import random
from numba import njit
import string
import time
import numpy as np


def main():

    list1Id = random.sample(range(1, 100000), 99000) #99000 losowych liczb
    list1Str = [''.join(random.choice(string.ascii_letters) for _ in range(3)) for _ in range(99000)] #99000 losowych 3 znakowych stringow
    list2Id = random.sample(range(1, 100000), 6000) #6000 losowych liczb
    list2Str = ["" for i in range(6000)] #lista ktora ma byc wypelniona danymi z list1Str
    list2Str_2 = ["" for i in range(6000)] #kopia list2Str dla uzycia set
    list2Str_3 = ["" for i in range(6000)] #kopia list2Str dla uzycia slownika

    print(list1Id[:10])
    print(list1Str[:10])
    print(list2Id[:10])
    print(list2Str[:10])
    print('------------------------------------------------------')


    #konwersja list pythonowych na tablice numpy (numba nie wspolpracuje z listami)
    list1Id_array = np.array(list1Id, dtype=np.int32)
    list2Id_array = np.array(list2Id, dtype=np.int32)
    list1Str_array = np.array(list1Str, dtype='<U3')
    list2Str_array = np.array(list2Str, dtype='<U3')


    #przeszukiwanie korzystajac ze zwyklych list
    start_time = time.time()
    for i in range(len(list2Id)):
        for j in range(len(list1Id)):
            if (list2Id[i] == list1Id[j]):
                list2Str[i] = list1Str[j]
                break
    print("--- %s seconds ---" % (time.time() - start_time), 'standardowe listy')


    #przeszukiwanie korzystajac z set
    start_time = time.time()
    list1IdSet = set(list1Id)
    for i in range(len(list2Id)):
        if list2Id[i] in list1IdSet:
            j = list1Id.index(list2Id[i])
            list2Str_2[i] = list1Str[j]
    print("--- %s seconds ---" % (time.time() - start_time), 'set')


    #przeszukiwanie korzystajac ze slownika
    start_time = time.time()
    list1Id_dict = {list1Id[i]: list1Str[i] for i in range(len(list1Id))}
    for i in range(len(list2Id)):
        if list2Id[i] in list1Id_dict:
            list2Str_3[i] = list1Id_dict[list2Id[i]]
    print("--- %s seconds ---" % (time.time() - start_time), 'slownik')


    #przeszukiwanie korzystajac z numba
    start_time = time.time()
    @njit
    def numbaFunction(list1Id_array, list2Id_array,list1Str_array,list2Str_array):
        for i in range(len(list2Id_array)):
            for j in range(len(list1Id_array)):
                if (list2Id_array[i] == list1Id_array[j]):
                    list2Str_array[i] = list1Str_array[j]
                    break
        return list2Str_array
    list2Str_array = numbaFunction(list1Id_array, list2Id_array,list1Str_array,list2Str_array)
    list2Str_4 = list(list2Str_array)
    print("--- %s seconds ---" % (time.time() - start_time), 'numba')


    print('------------------------------------------------------')
    print(list2Str[:10]) #wynik dla uzycia list
    print(list2Str_2[:10]) #wynik dla uzycia set
    print(list2Str_3[:10]) #wynik dla uzycia slownika
    print(list2Str_4[:10]) #wynik dla uzycia numba

if __name__ == "__main__":
    main()
