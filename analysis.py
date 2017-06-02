# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from hazm import *
import pyodbc
import matplotlib.pyplot as plt


user='user200'
password='safta313?'
database='TahlilDB'
port='1433'
TDS_Version='8.0'
server='B3H3SH1-PC'
driver='SQL Server'
con_string='DATABASE=%s;PORT=%s;TDS=%s;SERVER=%s;driver=%s' % (database,port,TDS_Version,server,driver)



orgsens='گوشی های خوبی هستند. راضیم ازش 3ماه میشه دارمش تا الان مشکلی پیش نیومده. تو این رنج قیمت گوشی خوبیه'
presens=orgsens
normalizer = Normalizer()
stemmer = Stemmer()
lemmatizer = Lemmatizer()

def CheckStopWord(word):
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql = "select ID from Tbl_PerStopWords where Word=?"
    result = cursor2.execute(sql, word)
    data = cursor2.fetchall()
    lent = len(data)
    cnxn2.commit()
    if(lent>0):
        return True
    else:
        return False

def GetWordValue(word):
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql = "select Value from Tbl_WordPer where Word=?"
    result = cursor2.execute(sql, word)
    data = cursor2.fetchall()
    lent = len(data)
    cnxn2.commit()
    if(lent>0):
        return (data[0][0])
    else:
        a=0
        return a


def GetWordSents(sent):
    wordsvalue=0
    words = word_tokenize(sent)
    lenwords = len(words)
    for ii in range(0, lenwords):
        # print(words[ii].strip())
        stemword=stemmer.stem(words[ii].strip())
        # print(stemword)
        # print(lemmatizer.lemmatize(words[ii].strip()))
        if ((CheckStopWord(stemword)==False) and (CheckStopWord(words[ii].strip())==False)):
            wv=GetWordValue(stemword)
            if wv==0:
                wv = GetWordValue(words[ii].strip())
            wordsvalue=wordsvalue+wv
    return wordsvalue

def GetSentensePer(doc):
    sentsvalue=0
    sents = sent_tokenize(doc)
    lensents = len(sents)
    for i in range(0, lensents):
        # print(sents[i])
        sv=GetWordSents(sents[i].strip())
        # print(sv)
        sentsvalue=sentsvalue+sv
    return sentsvalue

def AddDocPer(IDDoc,Value):
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql = "update DigiKalaReviews set PerValue=? where ID=?"
    cursor2.execute(sql, Value,IDDoc)
    cnxn2.commit()
    cnxn2.close()

def AddDocCountPer(IDCat,NegCount,PlusCount,NonCount):
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql = "update Tbl_DigiCategory set PlusValue=?,NegValue=?,NonValue=? where ID=?"
    cursor2.execute(sql, PlusCount,NegCount,NonCount,IDCat)
    cnxn2.commit()
    cnxn2.close()

def AnalysisDoc(idcat):
    maximumreviewvalue=0
    pluscount=0
    negcount=0
    noncount=0
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql ="SELECT * FROM DigiKalaReviews where Category=?"
    result = cursor2.execute(sql, idcat)
    data = cursor2.fetchall()
    lent = len(data)
    cnxn2.commit()
    for row in data:
         iddoc=row[0]
         review=row[2].strip()
         print(review)
         review=normalizer.normalize(review)
         reviewvalue = GetSentensePer(review)
         if(reviewvalue >0):
            pluscount=pluscount+1
         if(reviewvalue <0):
            negcount=negcount+1
         if (reviewvalue == 0):
             noncount=noncount+1
         maximumreviewvalue=maximumreviewvalue+reviewvalue
         AddDocPer(iddoc,reviewvalue)
         review=review.encode('utf8').decode('utf8')
         print(review)
         print(reviewvalue)
    AddDocCountPer(idcat,negcount,pluscount,noncount)
    cnxn2.close()

def NormalazePerValue(idcat):
    PerValueArrayPos = []
    PerValueArrayNeg = []
    PerValueArrayPosID = []
    PerValueArrayNegID = []
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql = "SELECT ID,PerValue FROM DigiKalaReviews where Category=?"
    result = cursor2.execute(sql, idcat)
    data = cursor2.fetchall()
    lent = len(data)
    cnxn2.commit()
    for row in data:
        id=row[0]
        val = row[1]
        if(val>0):
            PerValueArrayPos.append(val)
            PerValueArrayPosID.append(id)
        else:
            PerValueArrayNeg.append(val)
            PerValueArrayNegID.append(id)
    maxpos=max(PerValueArrayPos)
    minneg= min(PerValueArrayNeg)*-1
    index=0
    for v in PerValueArrayPos:
        normv=v/maxpos
        sql = "UPDATE DigiKalaReviews  set PerValueNorm=? where ID=?"
        cursor2.execute(sql,normv, PerValueArrayPosID[index])
        cnxn2.commit()
        index+=1
    index=0
    for v in PerValueArrayNeg:
        if(v==0):
            normv=0
        else:
            normv=v/minneg
        sql = "UPDATE DigiKalaReviews  set PerValueNorm=? where ID=?"
        cursor2.execute(sql,normv, PerValueArrayNegID[index])
        cnxn2.commit()
        index+=1
    cnxn2.close()

def SetMSEValue(id,mse):
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql = "update DigiKalaReviews set MSEValue=? where ID=?"
    cursor2.execute(sql, mse,id)
    cnxn2.commit()
    cnxn2.close()


def MSE(idcat):
    pervalue = 0
    mainvalue = 0
    msevalue=0
    summsevalue=0
    MainMSEValue = 0
    TotalRecordCount = 0
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql = "SELECT ID,PerValueNorm,MainValue FROM DigiKalaReviews where Category=? order by ID ASC"
    result = cursor2.execute(sql,idcat)
    data = cursor2.fetchall()
    lent = len(data)
    cnxn2.commit()
    for row in data:
        iddoc = row[0]
        pervalue=row[1]
        mainvalue=row[2]
        if mainvalue== '' or mainvalue == None:
            mainvalue=0
        if pervalue == '' or pervalue == None:
            pervalue = 0
        mainvalue=float(mainvalue*1)
        pervalue=float(pervalue*1)
        msevalue=mainvalue-pervalue
        summsevalue+=msevalue
        TotalRecordCount+=1
        SetMSEValue(iddoc,msevalue)
    cnxn2.close()
    MainMSEValue=summsevalue/TotalRecordCount
    return MainMSEValue

def Chart(idcat):
    poscount=0
    negcount=0
    noncount=0
    cnxn2 = pyodbc.connect(con_string)
    cursor2 = cnxn2.cursor()
    sql = "SELECT PerValue FROM DigiKalaReviews where Category=? order by ID ASC"
    result = cursor2.execute(sql, idcat)
    data = cursor2.fetchall()
    lent = len(data)
    cnxn2.commit()
    for row in data:
        pervalue = row[0]
        if pervalue >0 :
            poscount+=1
        elif pervalue <0:
            negcount+=1
        else:
            noncount+=1
    cnxn2.close()
    print(poscount,negcount,noncount)
    labels = 'Positive', 'Negative', 'Neutral'
    sizes = [poscount, negcount, noncount]
    colors = ['yellowgreen', 'lightcoral', 'lightskyblue']
    explode = (0.1, 0, 0)  # explode 1st slice
    # Plot
    plt.pie(sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=140)
    plt.axis('equal')
    plt.show()

# NormalazePerValue(104)
MSEValue=MSE(104)
print(MSEValue)
# Chart(104)

