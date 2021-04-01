import xlwt

data = open("student.txt",'r')
book = xlwt.Workbook()

style1 = xlwt.easyxf('align: wrap yes,vert centre, horiz left; pattern: pattern solid,fore-colour light_yellow;border: left thin,right thin,top thin,bottom thin')

for stud_data in data.readlines():
    student_data = stud_data.split(",")

    sheet = book.add_sheet(student_data[0])
    sheet.write(1,0,"Subject",style1)
    sheet.write(1,1,"Term1",style1)
    sheet.write(1,2,"Term2",style1)
    sheet.write(1,3,"Final",style1)

    sheet.write(2, 0, "Math",style1)
    sheet.write(2, 1,student_data[2])
    sheet.write(2, 2,student_data[8])
    sheet.write(2, 3,student_data[14])

    sheet.write(3, 0, "Science",style1)
    sheet.write(3, 1, student_data[3])
    sheet.write(3, 2, student_data[9])
    sheet.write(3, 3, student_data[15])

    sheet.write(4, 0, "Social Science",style1)
    sheet.write(4, 1, student_data[4])
    sheet.write(4, 2, student_data[10])
    sheet.write(4, 3, student_data[16])

    sheet.write(5, 0, "English",style1)
    sheet.write(5, 1, student_data[5])
    sheet.write(5, 2, student_data[11])
    sheet.write(5, 3, student_data[17])

    sheet.write(6, 0, "Hindi",style1)
    sheet.write(6, 1, student_data[6])
    sheet.write(6, 2, student_data[12])
    sheet.write(6, 3, student_data[18])

    sheet.write(7,0,"Total",style1)
    term1_Total = int(student_data[2])+int(student_data[3])+int(student_data[4])+int(student_data[5])+int(student_data[6])
    term2_Total = int(student_data[8] )+ int(student_data[9] )+ int(student_data[10]) + int(student_data[11] )+ int(student_data[12])
    term3_Total = int(student_data[14] )+ int(student_data[15]) + int(student_data[16]) + int(student_data[17]) + int(student_data[18])

    sheet.write(7,1,term1_Total)
    sheet.write(7,2,term2_Total)
    sheet.write(7,3,term3_Total)

    per_term1 = (term1_Total *100/500)
    per_term2 = (term2_Total *100/500)
    per_term3 = (term3_Total *100/500)
    per_Total = (term1_Total+term2_Total+term3_Total)*100/1500

    sheet.write(9,0,"Term1 Percentage",style1)
    sheet.write(10,0,per_term1)

    sheet.write(9,1,"Term2 Percentage",style1)
    sheet.write(10,1,per_term2)

    sheet.write(9,2,"Term3 Percentage",style1)
    sheet.write(10,2,per_term3)

    sheet.write(9,3,"Overall Perecentae",style1)
    sheet.write(10,3,per_Total)

    if(per_Total>=80):
        grd ="A"
    elif(per_Total>=70):
        grd ="B"
    elif(per_Total>=60):
        grd ="C"
    elif(per_Total>=50):
        grd ="D"
    else:
        grd ="Fail"

    sheet.write(11,0,"Grade",style1)
    sheet.write(12,0,grd)

book.save("student_report.xls")
