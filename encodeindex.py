# -*- coding: utf-8 -*-
"""
Created on Tue Nov 29 17:08:06 2016

@author: Phillip Dix
"""

def decodeIndex(a):
    b=0
    for i in range(0,len(a)):
        b+=(26**i)*(ord(a[len(a)-(i+1)])-64)
    return b
    
def encodeIndex(a):  
    if a <= 26:
        z=chr(a+64)
    else:
        remainder = a%26
        b=int(a/26)
        if remainder ==0:
            remainder=26
            b=b-1
        z=encodeIndex(b) + chr(remainder+64)
    return z
    
def decodeList(firstrow,lastrow,firstcol,lastcol):
    A='['
    for i in range(int(firstrow),int(lastrow)+1):
        for j in range(decodeIndex(firstcol)-1,decodeIndex(lastcol)):
            A=A+'wb[self.scr.edit.WBindex][self.scr.edit.WSindex]["'+str(encodeIndex(j))+str(i)+'"],'
    A = A[:-1]
    return A+']'
    
if __name__ == '__main__':
    print decodeList(5,11,'A','E')