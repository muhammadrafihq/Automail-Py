###### to do nya: https://myaccount.google.com/intro/security utk bikin less secure di gmail nya

import os
import sys
import modul_SP

print('YOGYA - Speed Report v1.1')
print ('(c) SAB 2021 - RAFI')
print ('---------------------------------------')
print ('')
print ('Changelog:')
print ('Jan-2022 : Tarikan Yogya ')
print ('')

# this code change dir to posisi dimana file kite berade
os.chdir(os.path.dirname(os.path.abspath(__file__)))

cwd=os.getcwd()
#print (cwd)
#testing

adakah=os.path.exists('./WN S-26 LP SAB Yogya Gebyar Hebat.xlsx')
if adakah==True :
    os.remove("WN S-26 LP SAB Yogya Gebyar Hebat.xlsx")
    print ('delete old data Yogya')

modul_SP.Product()
print ('End of Code')


