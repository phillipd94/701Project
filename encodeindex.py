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
#I think we should all just take a moment to notice how clever this is
    index='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    if a<26:
        return index[a]
    else:
        return encodeIndex(a//26) + index[a%26]

def decodeList(firstrow,lastrow,firstcol,lastcol):
    A='['
    for i in range(int(firstrow),int(lastrow)+1):
        for j in range(decodeIndex(firstcol)-1,decodeIndex(lastcol)):
            A=A+'wb[self.scr.edit.WBindex][self.scr.edit.WSindex]["'+str(encodeIndex(j))+str(i)+'"],'
    A = A[:-1]
    return A+']'
    
if __name__ == '__main__':
    print decodeList(5,11,'A','E')