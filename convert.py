#image to text
import pytesseract as pt
from PIL import Image
import re
import pdftotext
import xlsxwriter


# im=Image.open("statement1.jpg")
# text = pt.image_to_string(im, lang="eng")
# # print(text)

#Load your PDF
with open("SampleStatement.pdf", "rb") as f:
    pdf = pdftotext.PDF(f)

text=''
for page in pdf:
    text=text+str(page)

f=open("AccDetailsSample.txt", "w+")

workbook= xlsxwriter.Workbook('AccountDetails.xlsx')
worksheet=workbook.add_worksheet()
row=0
col=0

pattern='Account.Number.*[0-9]{11,11}|Account No.*\d+[A-Z]+\d+[A-Z]*[0-9]*-[A-Z]*[0-9]*|Account No.*[0-9]{11,11}'
res=re.findall(pattern, text)
if res:
    f.write(str(res[0])+"\n")
    worksheet.write(row, col, res[0])
    col += 1

pattern= 'Name...[A-Z]+'
res=re.findall(pattern, text)
if res:
    f.write(str(res[0])+"\n")
    worksheet.write(row,col, res[0])
    col+=1

pattern= 'Email.ID...\w+@\w+.\w{3,3}|Email.*'
res=re.findall(pattern, text)
if res:
    f.write(str(res[0])+"\n")
    worksheet.write(row, col, res[0])
    col += 1

#pattern= '.*IFSC.*' #IFSC
pattern= '[A-Z]{4,5}\d{6,7}'
res=re.findall(pattern, text)
if res:
    f.write("IFSC:"+str(res[0])+"\n")
    worksheet.write(row,col, res[0])
    col+=1

pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}|[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}.to.[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}'
res=re.findall(pattern, text)
if res:
    f.write('Period: '+str(res[0])+'\n')
    worksheet.write(row, col, res[0])
    col += 1

pattern= '.*Balance.{0,3}[0-9]{0,10}\W[0-9]{2,2}|.*balance.*|.*BALANCE.*|.*Balance.*'
res=re.findall(pattern, text)
f.write(str(res[0])+"\n")
worksheet.write(row,col, res[0])
col+=1

pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.{0,3}[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*|\d{2,2}\W\d{2,2}\W\d{4,4}.*[A-Z]+.*|[0-9]{2,2}[a-zA-Z]{3,3}[0-9]{2,2}.*|[0-9][0-9]\W[0-9][0-9]\W[0-9]{2,2}.*|[0-9][0-9]\W[a-zA-Z]{3,3}\W[0-9][0-9].*'
p1='January.*|February.*|March.*|April.*|May.*|June.*|July.*|August.*|September.*|October.*|November.*|December.*'
p2= 'Jan.*|Feb.*|Mar.*|Apr.*|May.*|Jun.*|Jul.*|Aug.*|Sep.*|Oct.*|Nov.*|Dec.*'
res=re.findall(p2, text)
f.write('Transactions:\n')
for strs in res[3:]:
    f.write(str(strs)+'\n')
    worksheet.write(row, col, strs)
    col += 1

#reading acc details
f=open("AccDetailsSample.txt", "r")
f1=f.readlines()
for x in f1:
    print(x)


##########################################################

# with open("statement2.pdf", "rb") as f:
#     pdf = pdftotext.PDF(f)
#
# text=''
# for page in pdf:
#     text=text+str(page)
#
# row=1
# col=0
#
# pattern='Account.Number.*[0-9]{11,11}|Account No.*\d+[A-Z]+\d+[A-Z]*[0-9]*-[A-Z]*[0-9]*|Account No.*[0-9]{11,11}'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= 'Name...[A-Z]+'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row,col, res[0])
# col+=1
#
# pattern= 'Email.ID...\w+@\w+.\w{3,3}|Email.*'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# #pattern= '.*IFSC.*' #IFSC
# pattern= '[A-Z]{4,5}\d{6,7}'
# res=re.findall(pattern, text)
# #if !res[0]:
# #f.write("IFSC:"+str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}|[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}.to.[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}'
# res=re.findall(pattern, text)
# if res:
# #    f.write('Period: '+str(res[0])+'\n')
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= '.*Balance.{0,3}[0-9]{0,10}\W[0-9]{2,2}|.*balance.*|.*BALANCE.*'
# res=re.findall(pattern, text)
# #f.write(str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.{0,3}[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*|\d{2,2}\W\d{2,2}\W\d{4,4}.*[A-Z]+.*|[0-9]{2,2}[a-zA-Z]{3,3}[0-9]{2,2}.*|[0-9][0-9]\W[0-9][0-9]\W[0-9]{2,2}.*'
#
# res=re.findall(pattern, text)
# #f.write('Transactions:\n')
# for strs in res:
#     #f.write(str(strs)+'\n')
#     worksheet.write(row, col, strs)
#     col += 1
#
#
#
# ##########################################################
#
# with open("statement3.pdf", "rb") as f:
#     pdf = pdftotext.PDF(f)
#
# text=''
# for page in pdf:
#     text=text+str(page)
#
# row=2
# col=0
#
# pattern='Account.Number.*[0-9]{11,11}|Account No.*\d+[A-Z]+\d+[A-Z]*[0-9]*-[A-Z]*[0-9]*|Account No.*[0-9]{11,11}'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= 'Name...[A-Z]+'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row,col, res[0])
# col+=1
#
# pattern= 'Email.ID...\w+@\w+.\w{3,3}|Email.*'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# #pattern= '.*IFSC.*' #IFSC
# pattern= '[A-Z]{4,5}\d{6,7}'
# res=re.findall(pattern, text)
# #if !res[0]:
# #f.write("IFSC:"+str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}|[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}.to.[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}'
# res=re.findall(pattern, text)
# if res:
# #    f.write('Period: '+str(res[0])+'\n')
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= '.*Balance.{0,3}[0-9]{0,10}\W[0-9]{2,2}|.*balance.*|.*BALANCE.*'
# res=re.findall(pattern, text)
# #f.write(str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.{0,3}[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*|\d{2,2}\W\d{2,2}\W\d{4,4}.*[A-Z]+.*|[0-9]{2,2}[a-zA-Z]{3,3}[0-9]{2,2}.*|[0-9][0-9]\W[0-9][0-9]\W[0-9]{2,2}.*'
#
# res=re.findall(pattern, text)
# #f.write('Transactions:\n')
# for strs in res:
#     #f.write(str(strs)+'\n')
#     worksheet.write(row, col, strs)
#     col += 1
#
#

#
# ##########################################################
# with open("statement4.pdf", "rb") as f:
#     pdf = pdftotext.PDF(f)
#
# text=''
# for page in pdf:
#     text=text+str(page)
#
# row=3
# col=0
#
# pattern='Account.Number.*[0-9]{11,11}|Account No.*\d+[A-Z]+\d+[A-Z]*[0-9]*-[A-Z]*[0-9]*|Account No.*[0-9]{11,11}'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= 'Name...[A-Z]+'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row,col, res[0])
# col+=1
#
# pattern= 'Email.ID...\w+@\w+.\w{3,3}|Email.*'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# #pattern= '.*IFSC.*' #IFSC
# pattern= '[A-Z]{4,5}\d{6,7}'
# res=re.findall(pattern, text)
# #if !res[0]:
# #f.write("IFSC:"+str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}|[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}.to.[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}'
# res=re.findall(pattern, text)
# if res:
# #    f.write('Period: '+str(res[0])+'\n')
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= '.*Balance.{0,3}[0-9]{0,10}\W[0-9]{2,2}|.*balance.*|.*BALANCE.*'
# res=re.findall(pattern, text)
# #f.write(str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.{0,3}[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*|\d{2,2}\W\d{2,2}\W\d{4,4}.*[A-Z]+.*|[0-9]{2,2}[a-zA-Z]{3,3}[0-9]{2,2}.*|[0-9][0-9]\W[0-9][0-9]\W[0-9]{2,2}.*'
#
# res=re.findall(pattern, text)
# #f.write('Transactions:\n')
# for strs in res:
#     #f.write(str(strs)+'\n')
#     worksheet.write(row, col, strs)
#     col += 1

# #######################################
#
# im=Image.open("s42.jpg")
# text = pt.image_to_string(im, lang="eng")
# # print(text)
#
# #Load your PDF
# # with open("statement2.pdf", "rb") as f:
# #     pdf = pdftotext.PDF(f)
# #
# # text=''
# # for page in pdf:
# #     text=text+str(page)
#
# #f=open("AccDetails2.txt", "w+")
#
# row=4
# col=0
#
# pattern='Account.Number.*[0-9]{11,11}|Account No.*\d+[A-Z]+\d+[A-Z]*[0-9]*-[A-Z]*[0-9]*|Account No.*[0-9]{11,11}'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= 'Name...[A-Z]+'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row,col, res[0])
# col+=1
#
# pattern= 'Email.ID...\w+@\w+.\w{3,3}|Email.*'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# #pattern= '.*IFSC.*' #IFSC
# pattern= '[A-Z]{4,5}\d{6,7}'
# res=re.findall(pattern, text)
# #if !res[0]:
# #f.write("IFSC:"+str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}|[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}.to.[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}'
# res=re.findall(pattern, text)
# if res:
# #    f.write('Period: '+str(res[0])+'\n')
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= '.*Balance.{0,3}[0-9]{0,10}\W[0-9]{2,2}|.*balance.*|.*BALANCE.*'
# res=re.findall(pattern, text)
# #f.write(str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.{0,3}[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*|\d{2,2}\W\d{2,2}\W\d{4,4}.*[A-Z]+.*|[0-9]{2,2}[a-zA-Z]{3,3}[0-9]{2,2}.*|[0-9][0-9]\W[0-9][0-9]\W[0-9]{2,2}.*'
#
# res=re.findall(pattern, text)
# #f.write('Transactions:\n')
# for strs in res:
#     #f.write(str(strs)+'\n')
#     worksheet.write(row, col, strs)
#     col += 1
#
#
#
# ##########################################################
# with open("statement6.pdf", "rb") as f:
#     pdf = pdftotext.PDF(f)
#
# text=''
# for page in pdf:
#     text=text+str(page)
#
# row=5
# col=0
#
# pattern='Account.Number.*[0-9]{11,11}|Account No.*\d+[A-Z]+\d+[A-Z]*[0-9]*-[A-Z]*[0-9]*|Account No.*[0-9]{11,11}'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= 'Name...[A-Z]+'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row,col, res[0])
# col+=1
#
# pattern= 'Email.ID...\w+@\w+.\w{3,3}|Email.*'
# res=re.findall(pattern, text)
# if res:
# #    f.write(str(res[0])+"\n")
#     worksheet.write(row, col, res[0])
# col += 1
#
# #pattern= '.*IFSC.*' #IFSC
# pattern= '[A-Z]{4,5}\d{6,7}'
# res=re.findall(pattern, text)
# #if !res[0]:
# #f.write("IFSC:"+str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}|[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}.to.[A-Z][a-z]{2,2}.[0-9]{1,2}.*[0-9]{4,4}'
# res=re.findall(pattern, text)
# if res:
# #    f.write('Period: '+str(res[0])+'\n')
#     worksheet.write(row, col, res[0])
# col += 1
#
# pattern= '.*Balance.{0,3}[0-9]{0,10}\W[0-9]{2,2}|.*balance.*|.*BALANCE.*'
# res=re.findall(pattern, text)
# #f.write(str(res[0])+"\n")
# worksheet.write(row,col, res[0])
# col+=1
#
# pattern= '[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.{0,3}[0-9][0-9]\W[0-9][0-9]\W[0-9]{4,4}.*|\d{2,2}\W\d{2,2}\W\d{4,4}.*[A-Z]+.*|[0-9]{2,2}[a-zA-Z]{3,3}[0-9]{2,2}.*|[0-9][0-9]\W[0-9][0-9]\W[0-9]{2,2}.*'
#
# res=re.findall(pattern, text)
# #f.write('Transactions:\n')
# for strs in res:
#     #f.write(str(strs)+'\n')
#     worksheet.write(row, col, strs)
#     col += 1
#
# workbook.close()
#
# # #######################################
# #
# # workbook.close()
# # #f.close()
# #
# #
# # #reading acc details
# # # f=open("AccDetails2.txt", "r")
# # # f1=f.readlines()
# # # for x in f1:
# # #     print(x)
# #
