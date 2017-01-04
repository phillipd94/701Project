from encodeindex import decodeList
import re

class codeParser:

    def __init__(self):
        self.pattern=([r'[ \t\n\r\f\v+-=*()\[\]][A-Z]+\d+[ \t\n\r\f\v+-=*()\.]',
        r'[ \t\n\r\f\v+-=*()\[][A-Z]+\d+[ \t\n\r\f\v+-=*()\.]',
        r'=(.*wb\[self.scr.edit.WBindex\]\[self.scr.edit.WSindex\]\["[A-Z]+\d+"\].*)+',
        r'\[\s*[A-Z]+\d+\s*-\s*[A-Z]+\d+\s*\]',
        r'\snewWorkbook\(["\'].+["\']\)\s',
        r'\ssetActiveWorkbook\(\d+\)\s',
        r'\snewWorksheet\(["\'].+["\']\)\s',
        r'\ssetActiveWorksheet\(\d+\)\s',
        r'[ \t\n\r\f\1v+-=*()<>][A-Z]+\d+[ \t\n\r\f\v+-=*()\.><]',
        r'\[\s*[A-Z]+\d+\s*-\s*[A-Z]+\d+\s*\]'])
		
		
		
    def getCode(self,script):
        script='\n'+script+'\n'
        re.MULTILINE=True
        script=re.sub(self.pattern[3],self.rep3,script)        
        while (re.search(self.pattern[1],script)!=None): #capture overlapping instances of the self.pattern
            script=re.sub(self.pattern[1],self.rep1,script)
        while (re.search(self.pattern[0],script)!=None): #capture overlapping instances of the self.pattern
            script=re.sub(self.pattern[0],self.rep0,script)
            
        if "using cells as value" in script:
            script=re.sub(self.pattern[2],self.rep2,script)
            script=script.replace("using cells as value","\n")
        script=re.sub(self.pattern[4],self.rep4,script)
        script=re.sub(self.pattern[5],self.rep5,script)
        script=re.sub(self.pattern[6],self.rep6,script)
        script=re.sub(self.pattern[7],self.rep7,script)
        a='wb=self.scr.edit.activeWS\n'
        script=a+script+"\nself.locals=locals()"
        return script
        
        
    def rep7(self,a):
        a=a.group()
        z=re.findall(r'\s',a)
        b=re.search(r'\d+',a)
        c=z[0]+"self.scr.edit.WSindex="+b.group()+z[1]
        return c
		
    def rep6(self,a):
        a=a.group()
        z=re.findall(r'\s',a)
        b=re.search(r'[\'"].+[\'"]',a)
        c="wb=self.addWS("+b.group()+")"
        return z[0]+c+z[1]
		
    def rep5(self,a):
        a=a.group()
        z=re.findall(r'\s',a)
        b=re.search(r'\d+',a)
        c=z[0]+"self.scr.edit.WBindex="+b.group()+z[1]  #note that this DOES NOT change the active worksheet, possibly leading to an out of bounds error if my user is not careful
        return c
        
    def rep4(self,a):
        a=a.group()
        z=re.findall(r'\s',a)
        b=re.search(r'[\'"].+[\'"]',a)
        c="wb=self.addWB("+b.group()+")"
        return z[0]+c+z[1]

    def rep0(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'[A-Z]+\d+',self.rep1_0,a)
        return a
    def rep1_0(self,matchobj1):
        a=matchobj1.group()
        rep='["'+a+'"]'
        return rep
        
    def rep1(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'[A-Z]+\d+',self.rep1_1,a)
        return a
		
    def rep1_1(self,matchobj1):
        a=matchobj1.group()
        rep='wb[self.scr.edit.WBindex][self.scr.edit.WSindex]["'+a+'"]'
        return rep
    def rep2(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'wb\[self.scr.edit.WBindex\]\[self.scr.edit.WSindex\]\["[A-Z]+\d+"\]',self.rep2_2,a)
        return a
		
    def rep2_2(self,matchobj1):
        a=matchobj1.group()
        rep=a+'.value'
        return rep
		
    def rep3(self,b):
        a=b.group()
        a=a.replace(' ','')
        c=re.findall(r'\d+',a)
        d=re.findall(r'[A-Z]+',a)
        for i in range(0,len(c)):
            c[i]=int(c[i])
        c.sort()
        d.sort()
        return decodeList(c[0],c[1],d[0],d[1])
		
		
		
		
    def getInput(self,script):
    #         """method for parsing and colorizing keywords for cell names, input must be HTML"""
        script ='\n'+script+'\n'
        self.pattern1=r'[ \t\n\r\f\1v+-=*()<>][A-Z]+\d+[ \t\n\r\f\v+-=*()\.><]'
        self.pattern3=r'\[\s*[A-Z]+\d+\s*-\s*[A-Z]+\d+\s*\]'
        re.MULTILINE=True
        script=re.sub(self.pattern[9],self.crep3,script)        
        script=re.sub(self.pattern[8],self.crep1,script)
        valstate="using cells as value"
        script=script.replace(valstate,'<span style="color:blue;">'+valstate+'</span>')
        script=script.replace("np.",'<span style="color:blue;">'+"np."+'</span>')
        script=script.replace("xl.",'<span style="color:blue;">'+"xl."+'</span>')
        script=script.replace("setActiveWorksheet",'<span style="color:blue;">'+"setActiveWorksheet"+'</span>')
        script=script.replace("setActiveWorbook",'<span style="color:blue;">'+"setActiveWorkbook"+'</span>')
        script=script.replace("newWorksheet",'<span style="color:blue;">'+"newWorksheet"+'</span>')
        script=script.replace("newWorkbook",'<span style="color:blue;">'+"newWorkbook"+'</span>')
        script=script.replace('<span style=" color:#0000ff;">',"")
        
        return script
        
        
    
    def crep1(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'[A-Z]+\d+',self.crep1_1,a)
        return a
        
    def crep1_1(self,matchobj1):
        a=matchobj1.group()
        a='<span style="color:blue;">'+a+'</span>'
        return a
        
    #    def crep2(self,matchobj):
    #        a=matchobj.group()
    #        a=re.sub(r'wb.active\["[A-Z]+\d+"\]',self.crep2_2,a)
    #        #need to implement an easy user command to change active ws
#        return a
#    def crep2_2(self,matchobj1):
#        a=matchobj1.group()
#        rep=a+'.value'
#        return rep
    def crep3(self,b):
        a=b.group()
        a='<span style="color:blue;">'+a+'</span>'
        return a
		
		
		
		
		
if __name__ == '__main__':
    A=codeParser()
    f=open('testscript.txt','r')
    script1=f.read()
    f.close()
    print A.getCode(script1)
    print A.getInput(script1)
    print dir(A)