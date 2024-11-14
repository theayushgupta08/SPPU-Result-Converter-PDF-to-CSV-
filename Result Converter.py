import PyPDF2 as pp
import openpyxl as op
import csv
#import nltk


file_name = 'result.pdf'
reader = pp.PdfReader(file_name)
total_pages=(len(reader.pages)-1)


start=0
end=total_pages-1
header1=5
footer1=0
footer2=0


unwanted = ['',':']
mark_parameter = ['COURSE CODE','COURSE NAME','ISE','ESE','TOTAL','TW','PR','OR','Tot%','Crd','Grd','GP','CP','P&R','ORD'] #TUT
# Headers
student_detail_label = ['Seat Number', 'Student Name', 'Mother Name', 'PRN'] #Clg
student_credit_label = ['SGPA', 'Total Credit']  #Year
all_subjects = []
flag=0

headerlist = (['Seat','Student Name','Mother\'s Name','PRN'] + ['410241 DESIGN & ANALYSIS OF ALGO.'] + [''] * 13 
+ ['410242 MACHINE LEARNING'] + [''] * 13
+ ['410243 BLOCKCHAIN TECHNOLOGY '] + [''] * 13
+ ['410244D OBJ. ORIENTED MODL. & DESG. '] + [''] * 13
+ ['410245D SOFT. TEST. & QLTY ASSURANCE'] + [''] * 13
+ ['410246 LABORATORY PRACTICE - III '] + [''] * 13
+ ['410247 LABORATORY PRACTICE - IV'] + [''] * 13
+ ['410248 PROJECT STAGE - I'] + [''] * 13
+ ['410249B ENTERPRENEURSHIP DEVELOPMENT '] + [''] * 13
+ ['410250 HIGH PERFORMANCE COMPUTING '] + [''] * 13
+ ['410251 DEEP LEARNING '] + [''] * 13
+ [' 410252A NATURAL LANGUAGE PROCESSING '] + [''] * 13
+ [' 410253C BUSINESS INTELLIGENCE '] + [''] * 13
+ [' 410254 LABORATORY PRACTICE - V '] + [''] * 13
+ [' 410255 LABORATORY PRACTICE - VI  '] + [''] * 13
+ [' 410256 PROJECT STAGE II '] + [''] * 13
+ ['410257C SOCIAL MEDIA AND ANALYTICS '] + [''] * 13
+ [' 410501 HON-MACH. LEARN.& DATA SCI.  '] + [''] * 13
+ [' 410501 HON-MACH. LEARN.& DATA SCI. (Practical) '] + [''] * 13
+ [' 410503 HON-A.I. FOR BIG DATA ANA.  '] + [''] * 13
+ [' 410503 HON-A.I. FOR BIG DATA ANA. (Practical) '] + [''] * 13)


custom_list = [''] * 4 + ['Insem', 'Ensem', 'Total', 'TW', 'PR', 'OR', 'TUT', 'Tot%', 'Crd', 'Grd', 'GP', 'CP', 'P&R', 'ORD'] * 21

with open("Student_Details.csv","w",newline='') as csvfile:
    writer=csv.writer(csvfile)
    writer.writerow(headerlist)
    writer.writerow(custom_list)

for page in range(start,end+1):
    if flag==1:
        flag=0
        continue
    current_page=reader.pages[page]  #Pointing the Page
    data=current_page.extract_text()   #Extracting the text and storing into data
    data_1=data.split('\n') #Seperated data by new line  Stored in form of list
    size=len(data_1)   #Length of the list (no. of lines)
    if (page<end-1 and (data_1[size-1].startswith(".")==False)):
        page+=1
        next_page=reader.pages[page]
        data_=next_page.extract_text()
        data_1=data_1 + data_.split('\n')
        flag=1

    size=len(data_1)
    footer2=size-1

    for ind in range(0,size):                                         #indexing
        if (data_1[ind].startswith(".") and ind>4):
            footer1=ind
            break


    if page==0:      #for 1st page
        data_2_0=data_1[header1:footer1]     
        data_2_1=data_1[footer1+2:footer2]

    elif page<=end and footer2!=footer1:          #for rest of the page
        data_2_0=data_1[header1+1:footer1]
        data_2_1=data_1[footer1+2:footer2]
        
    elif page<=end and footer2==footer1:       #for the last page
        data_2_0 = data_1[header1+1:footer1]

    if data_1[footer1-3].startswith("FOURTH"):
        length_adjuster_1 = 3
        a1 = data_1[footer1 - 3]
        b1 = data_1[footer1 - 2]
        c1 = data_1[footer1 - 1]
    else :
        length_adjuster_1 = 1
        a1=b1=c1="Fail"
    if data_1[footer2-3].startswith("FOURTH") :
        length_adjuster_2 = 3
        a2 = data_1[footer2 - 3]
        b2 = data_1[footer2 - 2]
        c2 = data_1[footer2 - 1]
    else:
        length_adjuster_2 = 1
        a2=b2=c2="Fail"

    student_detail_1=data_2_0[0].split(" ")
    student_1 = [word for word in student_detail_1 if word not in unwanted]
    seat_index_1 = student_1.index('NO.:') + 1
    seat_1 = student_1[seat_index_1]
    mother_index_1 = student_1.index('MOTHER')+1
    mother_1 = student_1[mother_index_1]
    prn_index_1 = student_1.index('CLG.:') - 1
    prn_1 = student_1[prn_index_1]
    prn_1 = prn_1.lstrip(':')
    student_name_1 = student_1[seat_index_1 + 2:mother_index_1 - 1]
    student_info_1 = [seat_1, student_name_1, mother_1, prn_1]
    #print(page,data_2_0 )
    sem_ind_1 = data_2_0.index(' ')
    del data_2_0[sem_ind_1]
    # print(seat_1,seat_2,mother_1,mother_2,prn_1,prn_2)
    #print(student_name_1)
    # print(student_1[2])

    if page<end:
        student_detail_2=data_2_1[0].split(" ")
        student_2=[word for word in student_detail_2 if word not in unwanted]
        seat_index_2 = student_2.index('NO.:') + 1
        seat_2 = student_2[seat_index_2]
        mother_index_2 = student_2.index('MOTHER')+1
        mother_2 = student_2[mother_index_2]
        prn_index_2 = student_2.index('CLG.:')-1
        prn_2 = student_2[prn_index_2]
        prn_2 = prn_2.lstrip(':')
        student_name_2 = student_2[seat_index_2+2:mother_index_2-1]
        #print(student_2[2])
        student_info_2 = [seat_2,student_name_2,mother_2,prn_2]
        sem_ind_2 = data_2_1.index(' ')
        del data_2_1[sem_ind_2]

    all_subject_marks = []
    all_subject_marks_2 = []
    #print(data_2_0[3])


    for sub_num in range(2,(len(data_2_0)-length_adjuster_1)):
        sub_marks = data_2_0[sub_num]
        sub_marks = sub_marks.split(' ')
        sub_marks = [word for word in sub_marks if word not in unwanted]
        subject_name = sub_marks[1:-14]
        score = sub_marks[-14:]
        #print(score)  #std1 marks
        all_subject_marks.append(sub_marks[0])
        all_subject_marks.append(subject_name)
        for i in range(len(score)):
            score_builder = score[i]
            all_subject_marks.append(score_builder)
    if page<end:
        for sub_num2 in range(2,(len(data_2_1)-length_adjuster_2)):
            sub_marks2 = data_2_1[sub_num2]
            sub_marks2 = sub_marks2.split(' ')
            sub_marks2 = [word for word in sub_marks2 if word not in unwanted]
            subject_name2 = sub_marks2[1:-14]
            score2 = sub_marks2[-14:]
            #print(score2)   #std2 marks
            all_subject_marks_2.append(sub_marks2[0])
            all_subject_marks_2.append(subject_name2)
            for i in range(len(score2)):
                score_builder2 = score2[i]
                all_subject_marks_2.append(score_builder2)

    #print(student_info_1)
    #print(all_subject_marks)
    #print(all_subject_marks[all_subject_marks.index('410256')])
    std_details_1={
        'Seat No. ': student_info_1[0], 'Student Name ':' '.join(student_info_1[1]), "Mother's Name ": student_info_1[2],
        'PRN': student_info_1[3],
        #DAA
        'Insem': all_subject_marks[all_subject_marks.index('410241')+2][:3],
        'Ensem': all_subject_marks[all_subject_marks.index('410241')+3][:3],
        'Total': all_subject_marks[all_subject_marks.index('410241')+4][:3],
        'PW': all_subject_marks[all_subject_marks.index('410241')+5][:3],
        'PR': all_subject_marks[all_subject_marks.index('410241')+6][:3],
        'OR': all_subject_marks[all_subject_marks.index('410241')+7][:3],
        'TUT': all_subject_marks[all_subject_marks.index('410241')+8][:3],
        'TOT%': all_subject_marks[all_subject_marks.index('410241')+9][:3],
        'CRD': all_subject_marks[all_subject_marks.index('410241')+10][:3],
        'GRD': all_subject_marks[all_subject_marks.index('410241')+11][:3],
        'CP': all_subject_marks[all_subject_marks.index('410241')+12][:3],
        'GP': all_subject_marks[all_subject_marks.index('410241')+13][:3],
        'P&R': all_subject_marks[all_subject_marks.index('410241')+14][:3],
        'ORD': all_subject_marks[all_subject_marks.index('410241')+15][:3],
        #ML
        'Insem1': all_subject_marks[all_subject_marks.index('410242')+2][:3],
        'Ensem1': all_subject_marks[all_subject_marks.index('410242')+3][:3],
        'Total1': all_subject_marks[all_subject_marks.index('410242')+4][:3],
        'PW1': all_subject_marks[all_subject_marks.index('410242')+5][:3],
        'PR1': all_subject_marks[all_subject_marks.index('410242')+6][:3],
        'OR1': all_subject_marks[all_subject_marks.index('410242')+7][:3],
        'TUT1': all_subject_marks[all_subject_marks.index('410242')+8][:3],
        'TOT%1': all_subject_marks[all_subject_marks.index('410242')+9][:3],
        'CRD1': all_subject_marks[all_subject_marks.index('410242')+10][:3],
        'GRD1': all_subject_marks[all_subject_marks.index('410242')+11][:3],
        'CP1': all_subject_marks[all_subject_marks.index('410242')+12][:3],
        'GP1': all_subject_marks[all_subject_marks.index('410242')+13][:3],
        'P&R1': all_subject_marks[all_subject_marks.index('410242')+14][:3],
        'ORD1': all_subject_marks[all_subject_marks.index('410242')+15][:3],
        #BT
        'Insem2': all_subject_marks[all_subject_marks.index('410243')+2][:3],
        'Ensem2': all_subject_marks[all_subject_marks.index('410243')+3][:3],
        'Total2': all_subject_marks[all_subject_marks.index('410243')+4][:3],
        'PW2': all_subject_marks[all_subject_marks.index('410243')+5][:3],
        'PR2': all_subject_marks[all_subject_marks.index('410243')+6][:3],
        'OR2': all_subject_marks[all_subject_marks.index('410243')+7][:3],
        'TUT2': all_subject_marks[all_subject_marks.index('410243')+8][:3],
        'TOT%2': all_subject_marks[all_subject_marks.index('410243')+9][:3],
        'CRD2': all_subject_marks[all_subject_marks.index('410243')+10][:3],
        'GRD2': all_subject_marks[all_subject_marks.index('410243')+11][:3],
        'CP2': all_subject_marks[all_subject_marks.index('410243')+12][:3],
        'GP2': all_subject_marks[all_subject_marks.index('410243')+13][:3],
        'P&R2': all_subject_marks[all_subject_marks.index('410243')+14][:3],
        'ORD2': all_subject_marks[all_subject_marks.index('410243')+15][:3],
        #OOMD
        'Insem3': all_subject_marks[all_subject_marks.index('410244D')+2][:3],
        'Ensem3': all_subject_marks[all_subject_marks.index('410244D')+3][:3],
        'Total3': all_subject_marks[all_subject_marks.index('410244D')+4][:3],
        'PW3': all_subject_marks[all_subject_marks.index('410244D')+5][:3],
        'PR3': all_subject_marks[all_subject_marks.index('410244D')+6][:3],
        'OR3': all_subject_marks[all_subject_marks.index('410244D')+7][:3],
        'TUT3': all_subject_marks[all_subject_marks.index('410244D')+8][:3],
        'TOT%3': all_subject_marks[all_subject_marks.index('410244D')+9][:3],
        'CRD3': all_subject_marks[all_subject_marks.index('410244D')+10][:3],
        'GRD3': all_subject_marks[all_subject_marks.index('410244D')+11][:3],
        'CP3': all_subject_marks[all_subject_marks.index('410244D')+12][:3],
        'GP3': all_subject_marks[all_subject_marks.index('410244D')+13][:3],
        'P&R3': all_subject_marks[all_subject_marks.index('410244D')+14][:3],
        'ORD3': all_subject_marks[all_subject_marks.index('410244D')+15][:3],
        #STQA
        'Insem4': all_subject_marks[all_subject_marks.index('410245D')+2][:3],
        'Ensem4': all_subject_marks[all_subject_marks.index('410245D')+3][:3],
        'Total4': all_subject_marks[all_subject_marks.index('410245D')+4][:3],
        'PW4': all_subject_marks[all_subject_marks.index('410245D')+5][:3],
        'PR4': all_subject_marks[all_subject_marks.index('410245D')+6][:3],
        'OR4': all_subject_marks[all_subject_marks.index('410245D')+7][:3],
        'TUT4': all_subject_marks[all_subject_marks.index('410245D')+8][:3],
        'TOT%4': all_subject_marks[all_subject_marks.index('410245D')+9][:3],
        'CRD4': all_subject_marks[all_subject_marks.index('410245D')+10][:3],
        'GRD4': all_subject_marks[all_subject_marks.index('410245D')+11][:3],
        'CP4': all_subject_marks[all_subject_marks.index('410245D')+12][:3],
        'GP4': all_subject_marks[all_subject_marks.index('410245D')+13][:3],
        'P&R4': all_subject_marks[all_subject_marks.index('410245D')+14][:3],
        'ORD4': all_subject_marks[all_subject_marks.index('410245D')+15][:3],
        #LP3
        'Insem5': all_subject_marks[all_subject_marks.index('410246')+2][:3],
        'Ensem5': all_subject_marks[all_subject_marks.index('410246')+3][:3],
        'Total5': all_subject_marks[all_subject_marks.index('410246')+4][:3],
        'PW5': all_subject_marks[all_subject_marks.index('410246')+5][:3],
        'PR5': all_subject_marks[all_subject_marks.index('410246')+6][:3],
        'OR5': all_subject_marks[all_subject_marks.index('410246')+7][:3],
        'TUT5': all_subject_marks[all_subject_marks.index('410246')+8][:3],
        'TOT%5': all_subject_marks[all_subject_marks.index('410246')+9][:3],
        'CRD5': all_subject_marks[all_subject_marks.index('410246')+10][:3],
        'GRD5': all_subject_marks[all_subject_marks.index('410246')+11][:3],
        'CP5': all_subject_marks[all_subject_marks.index('410246')+12][:3],
        'GP5': all_subject_marks[all_subject_marks.index('410246')+13][:3],
        'P&R5': all_subject_marks[all_subject_marks.index('410246')+14][:3],
        'ORD5': all_subject_marks[all_subject_marks.index('410246')+15][:3],
        #LP4
        'Insem6': all_subject_marks[all_subject_marks.index('410247')+2][:3],
        'Ensem6': all_subject_marks[all_subject_marks.index('410247')+3][:3],
        'Total6': all_subject_marks[all_subject_marks.index('410247')+4][:3],
        'PW6': all_subject_marks[all_subject_marks.index('410247')+5][:3],
        'PR6': all_subject_marks[all_subject_marks.index('410247')+6][:3],
        'OR6': all_subject_marks[all_subject_marks.index('410247')+7][:3],
        'TUT6': all_subject_marks[all_subject_marks.index('410247')+8][:3],
        'TOT%6': all_subject_marks[all_subject_marks.index('410247')+9][:3],
        'CRD6': all_subject_marks[all_subject_marks.index('410247')+10][:3],
        'GRD6': all_subject_marks[all_subject_marks.index('410247')+11][:3],
        'CP6': all_subject_marks[all_subject_marks.index('410247')+12][:3],
        'GP6': all_subject_marks[all_subject_marks.index('410247')+13][:3],
        'P&R6': all_subject_marks[all_subject_marks.index('410247')+14][:3],
        'ORD6': all_subject_marks[all_subject_marks.index('410247')+15][:3],
        #PS1
        'Insem7': all_subject_marks[all_subject_marks.index('410248')+2][:3],
        'Ensem7': all_subject_marks[all_subject_marks.index('410248')+3][:3],
        'Total7': all_subject_marks[all_subject_marks.index('410248')+4][:3],
        'PW7': all_subject_marks[all_subject_marks.index('410248')+5][:3],
        'PR7': all_subject_marks[all_subject_marks.index('410248')+6][:3],
        'OR7': all_subject_marks[all_subject_marks.index('410248')+7][:3],
        'TUT7': all_subject_marks[all_subject_marks.index('410248')+8][:3],
        'TOT%7': all_subject_marks[all_subject_marks.index('410248')+9][:3],
        'CRD7': all_subject_marks[all_subject_marks.index('410248')+10][:3],
        'GRD7': all_subject_marks[all_subject_marks.index('410248')+11][:3],
        'CP7': all_subject_marks[all_subject_marks.index('410248')+12][:3],
        'GP7': all_subject_marks[all_subject_marks.index('410248')+13][:3],
        'P&R7': all_subject_marks[all_subject_marks.index('410248')+14][:3],
        'ORD7': all_subject_marks[all_subject_marks.index('410248')+15][:3],
        #ED
        'Insem8': all_subject_marks[all_subject_marks.index('410249B')+2][:3],
        'Ensem8': all_subject_marks[all_subject_marks.index('410249B')+3][:3],
        'Total8': all_subject_marks[all_subject_marks.index('410249B')+4][:3],
        'PW8': all_subject_marks[all_subject_marks.index('410249B')+5][:3],
        'PR8': all_subject_marks[all_subject_marks.index('410249B')+6][:3],
        'OR8': all_subject_marks[all_subject_marks.index('410249B')+7][:3],
        'TUT8': all_subject_marks[all_subject_marks.index('410249B')+8][:3],
        'TOT%8': all_subject_marks[all_subject_marks.index('410249B')+9][:3],
        'CRD8': all_subject_marks[all_subject_marks.index('410249B')+10][:3],
        'GRD8': all_subject_marks[all_subject_marks.index('410249B')+11][:3],
        'CP8': all_subject_marks[all_subject_marks.index('410249B')+12][:3],
        'GP8': all_subject_marks[all_subject_marks.index('410249B')+13][:3],
        'P&R8': all_subject_marks[all_subject_marks.index('410249B')+14][:3],
        'ORD8': all_subject_marks[all_subject_marks.index('410249B')+15][:3],
        #HPC
        'Insem9': all_subject_marks[all_subject_marks.index('410250')+2][:3],
        'Ensem9': all_subject_marks[all_subject_marks.index('410250')+3][:3],
        'Total9': all_subject_marks[all_subject_marks.index('410250')+4][:3],
        'PW9': all_subject_marks[all_subject_marks.index('410250')+5][:3],
        'PR9': all_subject_marks[all_subject_marks.index('410250')+6][:3],
        'OR9': all_subject_marks[all_subject_marks.index('410250')+7][:3],
        'TUT9': all_subject_marks[all_subject_marks.index('410250')+8][:3],
        'TOT%9': all_subject_marks[all_subject_marks.index('410250')+9][:3],
        'CRD9': all_subject_marks[all_subject_marks.index('410250')+10][:3],
        'GRD9': all_subject_marks[all_subject_marks.index('410250')+11][:3],
        'CP9': all_subject_marks[all_subject_marks.index('410250')+12][:3],
        'GP9': all_subject_marks[all_subject_marks.index('410250')+13][:3],
        'P&R9': all_subject_marks[all_subject_marks.index('410250')+14][:3],
        'ORD9': all_subject_marks[all_subject_marks.index('410250')+15][:3],
        #DL
        'Insem10': all_subject_marks[all_subject_marks.index('410251')+2][:3],
        'Ensem10': all_subject_marks[all_subject_marks.index('410251')+3][:3],
        'Total10': all_subject_marks[all_subject_marks.index('410251')+4][:3],
        'PW10': all_subject_marks[all_subject_marks.index('410251')+5][:3],
        'PR10': all_subject_marks[all_subject_marks.index('410251')+6][:3],
        'OR10': all_subject_marks[all_subject_marks.index('410251')+7][:3],
        'TUT10': all_subject_marks[all_subject_marks.index('410251')+8][:3],
        'TOT%10': all_subject_marks[all_subject_marks.index('410251')+9][:3],
        'CRD10': all_subject_marks[all_subject_marks.index('410251')+10][:3],
        'GRD10': all_subject_marks[all_subject_marks.index('410251')+11][:3],
        'CP10': all_subject_marks[all_subject_marks.index('410251')+12][:3],
        'GP10': all_subject_marks[all_subject_marks.index('410251')+13][:3],
        'P&R10': all_subject_marks[all_subject_marks.index('410251')+14][:3],
        'ORD10': all_subject_marks[all_subject_marks.index('410251')+15][:3],
        #NLP
        'Insem11': all_subject_marks[all_subject_marks.index('410252A')+2][:3],
        'Ensem11': all_subject_marks[all_subject_marks.index('410252A')+3][:3],
        'Total11': all_subject_marks[all_subject_marks.index('410252A')+4][:3],
        'PW11': all_subject_marks[all_subject_marks.index('410252A')+5][:3],
        'PR11': all_subject_marks[all_subject_marks.index('410252A')+6][:3],
        'OR11': all_subject_marks[all_subject_marks.index('410252A')+7][:3],
        'TUT11': all_subject_marks[all_subject_marks.index('410252A')+8][:3],
        'TOT%11': all_subject_marks[all_subject_marks.index('410252A')+9][:3],
        'CRD11': all_subject_marks[all_subject_marks.index('410252A')+10][:3],
        'GRD11': all_subject_marks[all_subject_marks.index('410252A')+11][:3],
        'CP11': all_subject_marks[all_subject_marks.index('410252A')+12][:3],
        'GP11': all_subject_marks[all_subject_marks.index('410252A')+13][:3],
        'P&R11': all_subject_marks[all_subject_marks.index('410252A')+14][:3],
        'ORD11': all_subject_marks[all_subject_marks.index('410252A')+15][:3],
        #BI
        'Insem12': all_subject_marks[all_subject_marks.index('410253C')+2][:3],
        'Ensem12': all_subject_marks[all_subject_marks.index('410253C')+3][:3],
        'Total12': all_subject_marks[all_subject_marks.index('410253C')+4][:3],
        'PW12': all_subject_marks[all_subject_marks.index('410253C')+5][:3],
        'PR12': all_subject_marks[all_subject_marks.index('410253C')+6][:3],
        'OR12': all_subject_marks[all_subject_marks.index('410253C')+7][:3],
        'TUT12': all_subject_marks[all_subject_marks.index('410253C')+8][:3],
        'TOT%12': all_subject_marks[all_subject_marks.index('410253C')+9][:3],
        'CRD12': all_subject_marks[all_subject_marks.index('410253C')+10][:3],
        'GRD12': all_subject_marks[all_subject_marks.index('410253C')+11][:3],
        'CP12': all_subject_marks[all_subject_marks.index('410253C')+12][:3],
        'GP12': all_subject_marks[all_subject_marks.index('410253C')+13][:3],
        'P&R12': all_subject_marks[all_subject_marks.index('410253C')+14][:3],
        'ORD12': all_subject_marks[all_subject_marks.index('410253C')+15][:3],
        #LP4
        'Insem13': all_subject_marks[all_subject_marks.index('410254')+2][:3],
        'Ensem13': all_subject_marks[all_subject_marks.index('410254')+3][:3],
        'Total13': all_subject_marks[all_subject_marks.index('410254')+4][:3],
        'PW13': all_subject_marks[all_subject_marks.index('410254')+5][:3],
        'PR13': all_subject_marks[all_subject_marks.index('410254')+6][:3],
        'OR13': all_subject_marks[all_subject_marks.index('410254')+7][:3],
        'TUT13': all_subject_marks[all_subject_marks.index('410254')+8][:3],
        'TOT%13': all_subject_marks[all_subject_marks.index('410254')+9][:3],
        'CRD13': all_subject_marks[all_subject_marks.index('410254')+10][:3],
        'GRD13': all_subject_marks[all_subject_marks.index('410254')+11][:3],
        'CP13': all_subject_marks[all_subject_marks.index('410254')+12][:3],
        'GP13': all_subject_marks[all_subject_marks.index('410254')+13][:3],
        'P&R13': all_subject_marks[all_subject_marks.index('410254')+14][:3],
        'ORD13': all_subject_marks[all_subject_marks.index('410254')+15][:3],
        #LP5
        'Insem14': all_subject_marks[all_subject_marks.index('410255')+2][:3],
        'Ensem14': all_subject_marks[all_subject_marks.index('410255')+3][:3],
        'Total14': all_subject_marks[all_subject_marks.index('410255')+4][:3],
        'PW14': all_subject_marks[all_subject_marks.index('410255')+5][:3],
        'PR14': all_subject_marks[all_subject_marks.index('410255')+6][:3],
        'OR14': all_subject_marks[all_subject_marks.index('410255')+7][:3],
        'TUT14': all_subject_marks[all_subject_marks.index('410255')+8][:3],
        'TOT%14': all_subject_marks[all_subject_marks.index('410255')+9][:3],
        'CRD14': all_subject_marks[all_subject_marks.index('410255')+10][:3],
        'GRD14': all_subject_marks[all_subject_marks.index('410255')+11][:3],
        'CP14': all_subject_marks[all_subject_marks.index('410255')+12][:3],
        'GP14': all_subject_marks[all_subject_marks.index('410255')+13][:3],
        'P&R14': all_subject_marks[all_subject_marks.index('410255')+14][:3],
        'ORD14': all_subject_marks[all_subject_marks.index('410255')+15][:3],
        #PS2
        'Insem15': all_subject_marks[all_subject_marks.index('410256') + 2][:3],
        'Ensem15': all_subject_marks[all_subject_marks.index('410256') + 3][:3],
        'Total15': all_subject_marks[all_subject_marks.index('410256') + 4][:3],
        'PW15': all_subject_marks[all_subject_marks.index('410256') + 5][:3],
        'PR15': all_subject_marks[all_subject_marks.index('410256') + 6][:3],
        'OR15': all_subject_marks[all_subject_marks.index('410256') + 7][:3],
        'TUT15': all_subject_marks[all_subject_marks.index('410256') + 8][:3],
        'TOT%15': all_subject_marks[all_subject_marks.index('410256') + 9][:3],
        'CRD15': all_subject_marks[all_subject_marks.index('410256') + 10][:3],
        'GRD15': all_subject_marks[all_subject_marks.index('410256') + 11][:3],
        'CP15': all_subject_marks[all_subject_marks.index('410256') + 12][:3],
        'GP15': all_subject_marks[all_subject_marks.index('410256') + 13][:3],
        'P&R15': all_subject_marks[all_subject_marks.index('410256') + 14][:3],
        'ORD15': all_subject_marks[all_subject_marks.index('410256') + 15][:3],
        #SMA
        'Insem16': all_subject_marks[all_subject_marks.index('410257C')+2][:3],
        'Ensem16': all_subject_marks[all_subject_marks.index('410257C')+3][:3],
        'Total16': all_subject_marks[all_subject_marks.index('410257C')+4][:3],
        'PW16': all_subject_marks[all_subject_marks.index('410257C')+5][:3],
        'PR16': all_subject_marks[all_subject_marks.index('410257C')+6][:3],
        'OR16': all_subject_marks[all_subject_marks.index('410257C')+7][:3],
        'TUT16': all_subject_marks[all_subject_marks.index('410257C')+8][:3],
        'TOT%16': all_subject_marks[all_subject_marks.index('410257C')+9][:3],
        'CRD16': all_subject_marks[all_subject_marks.index('410257C')+10][:3],
        'GRD16': all_subject_marks[all_subject_marks.index('410257C')+11][:3],
        'CP16': all_subject_marks[all_subject_marks.index('410257C')+12][:3],
        'GP16': all_subject_marks[all_subject_marks.index('410257C')+13][:3],
        'P&R16': all_subject_marks[all_subject_marks.index('410257C')+14][:3],
        'ORD16': all_subject_marks[all_subject_marks.index('410257C')+15][:3]
        
    }

    if page!=end:
        std_details_2={
        'Seat No. ': student_info_2[0], 'Student Name ':' '.join(student_info_2[1]), "Mother's Name ": student_info_2[2],
        'PRN': student_info_2[3],
        #DAA
        'Insem': all_subject_marks_2[all_subject_marks_2.index('410241')+2][:3],
        'Ensem': all_subject_marks_2[all_subject_marks_2.index('410241')+3][:3],
        'Total': all_subject_marks_2[all_subject_marks_2.index('410241')+4][:3],
        'PW': all_subject_marks_2[all_subject_marks_2.index('410241')+5][:3],
        'PR': all_subject_marks_2[all_subject_marks_2.index('410241')+6][:3],
        'OR': all_subject_marks_2[all_subject_marks_2.index('410241')+7][:3],
        'TUT': all_subject_marks_2[all_subject_marks_2.index('410241')+8][:3],
        'TOT%': all_subject_marks_2[all_subject_marks_2.index('410241')+9][:3],
        'CRD': all_subject_marks_2[all_subject_marks_2.index('410241')+10][:3],
        'GRD': all_subject_marks_2[all_subject_marks_2.index('410241')+11][:3],
        'CP': all_subject_marks_2[all_subject_marks_2.index('410241')+12][:3],
        'GP': all_subject_marks_2[all_subject_marks_2.index('410241')+13][:3],
        'P&R': all_subject_marks_2[all_subject_marks_2.index('410241')+14][:3],
        'ORD': all_subject_marks_2[all_subject_marks_2.index('410241')+15][:3],
        #ML
        'Insem1': all_subject_marks_2[all_subject_marks_2.index('410242')+2][:3],
        'Ensem1': all_subject_marks_2[all_subject_marks_2.index('410242')+3][:3],
        'Total1': all_subject_marks_2[all_subject_marks_2.index('410242')+4][:3],
        'PW1': all_subject_marks_2[all_subject_marks_2.index('410242')+5][:3],
        'PR1': all_subject_marks_2[all_subject_marks_2.index('410242')+6][:3],
        'OR1': all_subject_marks_2[all_subject_marks_2.index('410242')+7][:3],
        'TUT1': all_subject_marks_2[all_subject_marks_2.index('410242')+8][:3],
        'TOT%1': all_subject_marks_2[all_subject_marks_2.index('410242')+9][:3],
        'CRD1': all_subject_marks_2[all_subject_marks_2.index('410242')+10][:3],
        'GRD1': all_subject_marks_2[all_subject_marks_2.index('410242')+11][:3],
        'CP1': all_subject_marks_2[all_subject_marks_2.index('410242')+12][:3],
        'GP1': all_subject_marks_2[all_subject_marks_2.index('410242')+13][:3],
        'P&R1': all_subject_marks_2[all_subject_marks_2.index('410242')+14][:3],
        'ORD1': all_subject_marks_2[all_subject_marks_2.index('410242')+15][:3],
        #BT
        'Insem2': all_subject_marks_2[all_subject_marks_2.index('410243')+2][:3],
        'Ensem2': all_subject_marks_2[all_subject_marks_2.index('410243')+3][:3],
        'Total2': all_subject_marks_2[all_subject_marks_2.index('410243')+4][:3],
        'PW2': all_subject_marks_2[all_subject_marks_2.index('410243')+5][:3],
        'PR2': all_subject_marks_2[all_subject_marks_2.index('410243')+6][:3],
        'OR2': all_subject_marks_2[all_subject_marks_2.index('410243')+7][:3],
        'TUT2': all_subject_marks_2[all_subject_marks_2.index('410243')+8][:3],
        'TOT%2': all_subject_marks_2[all_subject_marks_2.index('410243')+9][:3],
        'CRD2': all_subject_marks_2[all_subject_marks_2.index('410243')+10][:3],
        'GRD2': all_subject_marks_2[all_subject_marks_2.index('410243')+11][:3],
        'CP2': all_subject_marks_2[all_subject_marks_2.index('410243')+12][:3],
        'GP2': all_subject_marks_2[all_subject_marks_2.index('410243')+13][:3],
        'P&R2': all_subject_marks_2[all_subject_marks_2.index('410243')+14][:3],
        'ORD2': all_subject_marks_2[all_subject_marks_2.index('410243')+15][:3],
        #OOMD
        'Insem3': all_subject_marks_2[all_subject_marks_2.index('410244D')+2][:3],
        'Ensem3': all_subject_marks_2[all_subject_marks_2.index('410244D')+3][:3],
        'Total3': all_subject_marks_2[all_subject_marks_2.index('410244D')+4][:3],
        'PW3': all_subject_marks_2[all_subject_marks_2.index('410244D')+5][:3],
        'PR3': all_subject_marks_2[all_subject_marks_2.index('410244D')+6][:3],
        'OR3': all_subject_marks_2[all_subject_marks_2.index('410244D')+7][:3],
        'TUT3': all_subject_marks_2[all_subject_marks_2.index('410244D')+8][:3],
        'TOT%3': all_subject_marks_2[all_subject_marks_2.index('410244D')+9][:3],
        'CRD3': all_subject_marks_2[all_subject_marks_2.index('410244D')+10][:3],
        'GRD3': all_subject_marks_2[all_subject_marks_2.index('410244D')+11][:3],
        'CP3': all_subject_marks_2[all_subject_marks_2.index('410244D')+12][:3],
        'GP3': all_subject_marks_2[all_subject_marks_2.index('410244D')+13][:3],
        'P&R3': all_subject_marks_2[all_subject_marks_2.index('410244D')+14][:3],
        'ORD3': all_subject_marks_2[all_subject_marks_2.index('410244D')+15][:3],
        #STQA
        'Insem4': all_subject_marks_2[all_subject_marks_2.index('410245D')+2][:3],
        'Ensem4': all_subject_marks_2[all_subject_marks_2.index('410245D')+3][:3],
        'Total4': all_subject_marks_2[all_subject_marks_2.index('410245D')+4][:3],
        'PW4': all_subject_marks_2[all_subject_marks_2.index('410245D')+5][:3],
        'PR4': all_subject_marks_2[all_subject_marks_2.index('410245D')+6][:3],
        'OR4': all_subject_marks_2[all_subject_marks_2.index('410245D')+7][:3],
        'TUT4': all_subject_marks_2[all_subject_marks_2.index('410245D')+8][:3],
        'TOT%4': all_subject_marks_2[all_subject_marks_2.index('410245D')+9][:3],
        'CRD4': all_subject_marks_2[all_subject_marks_2.index('410245D')+10][:3],
        'GRD4': all_subject_marks_2[all_subject_marks_2.index('410245D')+11][:3],
        'CP4': all_subject_marks_2[all_subject_marks_2.index('410245D')+12][:3],
        'GP4': all_subject_marks_2[all_subject_marks_2.index('410245D')+13][:3],
        'P&R4': all_subject_marks_2[all_subject_marks_2.index('410245D')+14][:3],
        'ORD4': all_subject_marks_2[all_subject_marks_2.index('410245D')+15][:3],
        #LP3
        'Insem5': all_subject_marks_2[all_subject_marks_2.index('410246')+2][:3],
        'Ensem5': all_subject_marks_2[all_subject_marks_2.index('410246')+3][:3],
        'Total5': all_subject_marks_2[all_subject_marks_2.index('410246')+4][:3],
        'PW5': all_subject_marks_2[all_subject_marks_2.index('410246')+5][:3],
        'PR5': all_subject_marks_2[all_subject_marks_2.index('410246')+6][:3],
        'OR5': all_subject_marks_2[all_subject_marks_2.index('410246')+7][:3],
        'TUT5': all_subject_marks_2[all_subject_marks_2.index('410246')+8][:3],
        'TOT%5': all_subject_marks_2[all_subject_marks_2.index('410246')+9][:3],
        'CRD5': all_subject_marks_2[all_subject_marks_2.index('410246')+10][:3],
        'GRD5': all_subject_marks_2[all_subject_marks_2.index('410246')+11][:3],
        'CP5': all_subject_marks_2[all_subject_marks_2.index('410246')+12][:3],
        'GP5': all_subject_marks_2[all_subject_marks_2.index('410246')+13][:3],
        'P&R5': all_subject_marks_2[all_subject_marks_2.index('410246')+14][:3],
        'ORD5': all_subject_marks_2[all_subject_marks_2.index('410246')+15][:3],
        #LP4
        'Insem6': all_subject_marks_2[all_subject_marks_2.index('410247')+2][:3],
        'Ensem6': all_subject_marks_2[all_subject_marks_2.index('410247')+3][:3],
        'Total6': all_subject_marks_2[all_subject_marks_2.index('410247')+4][:3],
        'PW6': all_subject_marks_2[all_subject_marks_2.index('410247')+5][:3],
        'PR6': all_subject_marks_2[all_subject_marks_2.index('410247')+6][:3],
        'OR6': all_subject_marks_2[all_subject_marks_2.index('410247')+7][:3],
        'TUT6': all_subject_marks_2[all_subject_marks_2.index('410247')+8][:3],
        'TOT%6': all_subject_marks_2[all_subject_marks_2.index('410247')+9][:3],
        'CRD6': all_subject_marks_2[all_subject_marks_2.index('410247')+10][:3],
        'GRD6': all_subject_marks_2[all_subject_marks_2.index('410247')+11][:3],
        'CP6': all_subject_marks_2[all_subject_marks_2.index('410247')+12][:3],
        'GP6': all_subject_marks_2[all_subject_marks_2.index('410247')+13][:3],
        'P&R6': all_subject_marks_2[all_subject_marks_2.index('410247')+14][:3],
        'ORD6': all_subject_marks_2[all_subject_marks_2.index('410247')+15][:3],
        #PS1
        'Insem7': all_subject_marks_2[all_subject_marks_2.index('410248')+2][:3],
        'Ensem7': all_subject_marks_2[all_subject_marks_2.index('410248')+3][:3],
        'Total7': all_subject_marks_2[all_subject_marks_2.index('410248')+4][:3],
        'PW7': all_subject_marks_2[all_subject_marks_2.index('410248')+5][:3],
        'PR7': all_subject_marks_2[all_subject_marks_2.index('410248')+6][:3],
        'OR7': all_subject_marks_2[all_subject_marks_2.index('410248')+7][:3],
        'TUT7': all_subject_marks_2[all_subject_marks_2.index('410248')+8][:3],
        'TOT%7': all_subject_marks_2[all_subject_marks_2.index('410248')+9][:3],
        'CRD7': all_subject_marks_2[all_subject_marks_2.index('410248')+10][:3],
        'GRD7': all_subject_marks_2[all_subject_marks_2.index('410248')+11][:3],
        'CP7': all_subject_marks_2[all_subject_marks_2.index('410248')+12][:3],
        'GP7': all_subject_marks_2[all_subject_marks_2.index('410248')+13][:3],
        'P&R7': all_subject_marks_2[all_subject_marks_2.index('410248')+14][:3],
        'ORD7': all_subject_marks_2[all_subject_marks_2.index('410248')+15][:3],
        #ED
        'Insem8': all_subject_marks_2[all_subject_marks_2.index('410249B')+2][:3],
        'Ensem8': all_subject_marks_2[all_subject_marks_2.index('410249B')+3][:3],
        'Total8': all_subject_marks_2[all_subject_marks_2.index('410249B')+4][:3],
        'PW8': all_subject_marks_2[all_subject_marks_2.index('410249B')+5][:3],
        'PR8': all_subject_marks_2[all_subject_marks_2.index('410249B')+6][:3],
        'OR8': all_subject_marks_2[all_subject_marks_2.index('410249B')+7][:3],
        'TUT8': all_subject_marks_2[all_subject_marks_2.index('410249B')+8][:3],
        'TOT%8': all_subject_marks_2[all_subject_marks_2.index('410249B')+9][:3],
        'CRD8': all_subject_marks_2[all_subject_marks_2.index('410249B')+10][:3],
        'GRD8': all_subject_marks_2[all_subject_marks_2.index('410249B')+11][:3],
        'CP8': all_subject_marks_2[all_subject_marks_2.index('410249B')+12][:3],
        'GP8': all_subject_marks_2[all_subject_marks_2.index('410249B')+13][:3],
        'P&R8': all_subject_marks_2[all_subject_marks_2.index('410249B')+14][:3],
        'ORD8': all_subject_marks_2[all_subject_marks_2.index('410249B')+15][:3],
        #HPC
        'Insem9': all_subject_marks_2[all_subject_marks_2.index('410250')+2][:3],
        'Ensem9': all_subject_marks_2[all_subject_marks_2.index('410250')+3][:3],
        'Total9': all_subject_marks_2[all_subject_marks_2.index('410250')+4][:3],
        'PW9': all_subject_marks_2[all_subject_marks_2.index('410250')+5][:3],
        'PR9': all_subject_marks_2[all_subject_marks_2.index('410250')+6][:3],
        'OR9': all_subject_marks_2[all_subject_marks_2.index('410250')+7][:3],
        'TUT9': all_subject_marks_2[all_subject_marks_2.index('410250')+8][:3],
        'TOT%9': all_subject_marks_2[all_subject_marks_2.index('410250')+9][:3],
        'CRD9': all_subject_marks_2[all_subject_marks_2.index('410250')+10][:3],
        'GRD9': all_subject_marks_2[all_subject_marks_2.index('410250')+11][:3],
        'CP9': all_subject_marks_2[all_subject_marks_2.index('410250')+12][:3],
        'GP9': all_subject_marks_2[all_subject_marks_2.index('410250')+13][:3],
        'P&R9': all_subject_marks_2[all_subject_marks_2.index('410250')+14][:3],
        'ORD9': all_subject_marks_2[all_subject_marks_2.index('410250')+15][:3],
        #DL
        'Insem10': all_subject_marks_2[all_subject_marks_2.index('410251')+2][:3],
        'Ensem10': all_subject_marks_2[all_subject_marks_2.index('410251')+3][:3],
        'Total10': all_subject_marks_2[all_subject_marks_2.index('410251')+4][:3],
        'PW10': all_subject_marks_2[all_subject_marks_2.index('410251')+5][:3],
        'PR10': all_subject_marks_2[all_subject_marks_2.index('410251')+6][:3],
        'OR10': all_subject_marks_2[all_subject_marks_2.index('410251')+7][:3],
        'TUT10': all_subject_marks_2[all_subject_marks_2.index('410251')+8][:3],
        'TOT%10': all_subject_marks_2[all_subject_marks_2.index('410251')+9][:3],
        'CRD10': all_subject_marks_2[all_subject_marks_2.index('410251')+10][:3],
        'GRD10': all_subject_marks_2[all_subject_marks_2.index('410251')+11][:3],
        'CP10': all_subject_marks_2[all_subject_marks_2.index('410251')+12][:3],
        'GP10': all_subject_marks_2[all_subject_marks_2.index('410251')+13][:3],
        'P&R10': all_subject_marks_2[all_subject_marks_2.index('410251')+14][:3],
        'ORD10': all_subject_marks_2[all_subject_marks_2.index('410251')+15][:3],
        #NLP
        'Insem11': all_subject_marks_2[all_subject_marks_2.index('410252A')+2][:3],
        'Ensem11': all_subject_marks_2[all_subject_marks_2.index('410252A')+3][:3],
        'Total11': all_subject_marks_2[all_subject_marks_2.index('410252A')+4][:3],
        'PW11': all_subject_marks_2[all_subject_marks_2.index('410252A')+5][:3],
        'PR11': all_subject_marks_2[all_subject_marks_2.index('410252A')+6][:3],
        'OR11': all_subject_marks_2[all_subject_marks_2.index('410252A')+7][:3],
        'TUT11': all_subject_marks_2[all_subject_marks_2.index('410252A')+8][:3],
        'TOT%11': all_subject_marks_2[all_subject_marks_2.index('410252A')+9][:3],
        'CRD11': all_subject_marks_2[all_subject_marks_2.index('410252A')+10][:3],
        'GRD11': all_subject_marks_2[all_subject_marks_2.index('410252A')+11][:3],
        'CP11': all_subject_marks_2[all_subject_marks_2.index('410252A')+12][:3],
        'GP11': all_subject_marks_2[all_subject_marks_2.index('410252A')+13][:3],
        'P&R11': all_subject_marks_2[all_subject_marks_2.index('410252A')+14][:3],
        'ORD11': all_subject_marks_2[all_subject_marks_2.index('410252A')+15][:3],
        #BI
        'Insem12': all_subject_marks_2[all_subject_marks_2.index('410253C')+2][:3],
        'Ensem12': all_subject_marks_2[all_subject_marks_2.index('410253C')+3][:3],
        'Total12': all_subject_marks_2[all_subject_marks_2.index('410253C')+4][:3],
        'PW12': all_subject_marks_2[all_subject_marks_2.index('410253C')+5][:3],
        'PR12': all_subject_marks_2[all_subject_marks_2.index('410253C')+6][:3],
        'OR12': all_subject_marks_2[all_subject_marks_2.index('410253C')+7][:3],
        'TUT12': all_subject_marks_2[all_subject_marks_2.index('410253C')+8][:3],
        'TOT%12': all_subject_marks_2[all_subject_marks_2.index('410253C')+9][:3],
        'CRD12': all_subject_marks_2[all_subject_marks_2.index('410253C')+10][:3],
        'GRD12': all_subject_marks_2[all_subject_marks_2.index('410253C')+11][:3],
        'CP12': all_subject_marks_2[all_subject_marks_2.index('410253C')+12][:3],
        'GP12': all_subject_marks_2[all_subject_marks_2.index('410253C')+13][:3],
        'P&R12': all_subject_marks_2[all_subject_marks_2.index('410253C')+14][:3],
        'ORD12': all_subject_marks_2[all_subject_marks_2.index('410253C')+15][:3],
        #LP4
        'Insem13': all_subject_marks_2[all_subject_marks_2.index('410254')+2][:3],
        'Ensem13': all_subject_marks_2[all_subject_marks_2.index('410254')+3][:3],
        'Total13': all_subject_marks_2[all_subject_marks_2.index('410254')+4][:3],
        'PW13': all_subject_marks_2[all_subject_marks_2.index('410254')+5][:3],
        'PR13': all_subject_marks_2[all_subject_marks_2.index('410254')+6][:3],
        'OR13': all_subject_marks_2[all_subject_marks_2.index('410254')+7][:3],
        'TUT13': all_subject_marks_2[all_subject_marks_2.index('410254')+8][:3],
        'TOT%13': all_subject_marks_2[all_subject_marks_2.index('410254')+9][:3],
        'CRD13': all_subject_marks_2[all_subject_marks_2.index('410254')+10][:3],
        'GRD13': all_subject_marks_2[all_subject_marks_2.index('410254')+11][:3],
        'CP13': all_subject_marks_2[all_subject_marks_2.index('410254')+12][:3],
        'GP13': all_subject_marks_2[all_subject_marks_2.index('410254')+13][:3],
        'P&R13': all_subject_marks_2[all_subject_marks_2.index('410254')+14][:3],
        'ORD13': all_subject_marks_2[all_subject_marks_2.index('410254')+15][:3],
        #LP5
        'Insem14': all_subject_marks_2[all_subject_marks_2.index('410255')+2][:3],
        'Ensem14': all_subject_marks_2[all_subject_marks_2.index('410255')+3][:3],
        'Total14': all_subject_marks_2[all_subject_marks_2.index('410255')+4][:3],
        'PW14': all_subject_marks_2[all_subject_marks_2.index('410255')+5][:3],
        'PR14': all_subject_marks_2[all_subject_marks_2.index('410255')+6][:3],
        'OR14': all_subject_marks_2[all_subject_marks_2.index('410255')+7][:3],
        'TUT14': all_subject_marks_2[all_subject_marks_2.index('410255')+8][:3],
        'TOT%14': all_subject_marks_2[all_subject_marks_2.index('410255')+9][:3],
        'CRD14': all_subject_marks_2[all_subject_marks_2.index('410255')+10][:3],
        'GRD14': all_subject_marks_2[all_subject_marks_2.index('410255')+11][:3],
        'CP14': all_subject_marks_2[all_subject_marks_2.index('410255')+12][:3],
        'GP14': all_subject_marks_2[all_subject_marks_2.index('410255')+13][:3],
        'P&R14': all_subject_marks_2[all_subject_marks_2.index('410255')+14][:3],
        'ORD14': all_subject_marks_2[all_subject_marks_2.index('410255')+15][:3],
        #PS2
        'Insem15': all_subject_marks_2[all_subject_marks_2.index('410256') + 2][:3],
        'Ensem15': all_subject_marks_2[all_subject_marks_2.index('410256') + 3][:3],
        'Total15': all_subject_marks_2[all_subject_marks_2.index('410256') + 4][:3],
        'PW15': all_subject_marks_2[all_subject_marks_2.index('410256') + 5][:3],
        'PR15': all_subject_marks_2[all_subject_marks_2.index('410256') + 6][:3],
        'OR15': all_subject_marks_2[all_subject_marks_2.index('410256') + 7][:3],
        'TUT15': all_subject_marks_2[all_subject_marks_2.index('410256') + 8][:3],
        'TOT%15': all_subject_marks_2[all_subject_marks_2.index('410256') + 9][:3],
        'CRD15': all_subject_marks_2[all_subject_marks_2.index('410256') + 10][:3],
        'GRD15': all_subject_marks_2[all_subject_marks_2.index('410256') + 11][:3],
        'CP15': all_subject_marks_2[all_subject_marks_2.index('410256') + 12][:3],
        'GP15': all_subject_marks_2[all_subject_marks_2.index('410256') + 13][:3],
        'P&R15': all_subject_marks_2[all_subject_marks_2.index('410256') + 14][:3],
        'ORD15': all_subject_marks_2[all_subject_marks_2.index('410256') + 15][:3],
        #SMA
        'Insem16': all_subject_marks_2[all_subject_marks_2.index('410257C')+2][:3],
        'Ensem16': all_subject_marks_2[all_subject_marks_2.index('410257C')+3][:3],
        'Total16': all_subject_marks_2[all_subject_marks_2.index('410257C')+4][:3],
        'PW16': all_subject_marks_2[all_subject_marks_2.index('410257C')+5][:3],
        'PR16': all_subject_marks_2[all_subject_marks_2.index('410257C')+6][:3],
        'OR16': all_subject_marks_2[all_subject_marks_2.index('410257C')+7][:3],
        'TUT16': all_subject_marks_2[all_subject_marks_2.index('410257C')+8][:3],
        'TOT%16': all_subject_marks_2[all_subject_marks_2.index('410257C')+9][:3],
        'CRD16': all_subject_marks_2[all_subject_marks_2.index('410257C')+10][:3],
        'GRD16': all_subject_marks_2[all_subject_marks_2.index('410257C')+11][:3],
        'CP16': all_subject_marks_2[all_subject_marks_2.index('410257C')+12][:3],
        'GP16': all_subject_marks_2[all_subject_marks_2.index('410257C')+13][:3],
        'P&R16': all_subject_marks_2[all_subject_marks_2.index('410257C')+14][:3],
        'ORD16': all_subject_marks_2[all_subject_marks_2.index('410257C')+15][:3]
    }
    
    #print(all_subject_marks_2[all_subject_marks_2.index('410501')+2][:3])
    #print(all_subject_marks[all_subject_marks.index('410501')+2][:3])

    #HON-MACH. LEARN.& DATA SCI for std1
    if '410501' in all_subject_marks:
        std_details_1['Insem17'] = all_subject_marks[all_subject_marks.index('410501')+2][:3]
        std_details_1['Ensem17'] = all_subject_marks[all_subject_marks.index('410501')+3][:3]
        std_details_1['Total17'] = all_subject_marks[all_subject_marks.index('410501')+4][:3]
        std_details_1['PW17'] = all_subject_marks[all_subject_marks.index('410501')+5][:3]
        std_details_1['PR17'] = all_subject_marks[all_subject_marks.index('410501')+6][:3]
        std_details_1['OR17'] = all_subject_marks[all_subject_marks.index('410501')+7][:3]
        std_details_1['TUT17'] = all_subject_marks[all_subject_marks.index('410501')+8][:3]
        std_details_1['TOT%17'] = all_subject_marks[all_subject_marks.index('410501')+9][:3]
        std_details_1['CRD17'] = all_subject_marks[all_subject_marks.index('410501')+10][:3]
        std_details_1['GRD17'] = all_subject_marks[all_subject_marks.index('410501')+11][:3]
        std_details_1['CP17'] = all_subject_marks[all_subject_marks.index('410501')+12][:3]
        std_details_1['GP17'] = all_subject_marks[all_subject_marks.index('410501')+13][:3]
        std_details_1['P&R17'] = all_subject_marks[all_subject_marks.index('410501')+14][:3]
        std_details_1['ORD17'] = all_subject_marks[all_subject_marks.index('410501')+15][:3]
        all_subject_marks.remove("410501")
        std_details_1['Insem18'] = all_subject_marks[all_subject_marks.index('410501')+2][:3]
        std_details_1['Ensem18'] = all_subject_marks[all_subject_marks.index('410501')+3][:3]
        std_details_1['Total18'] = all_subject_marks[all_subject_marks.index('410501')+4][:3]
        std_details_1['PW18'] = all_subject_marks[all_subject_marks.index('410501')+5][:3]
        std_details_1['PR18'] = all_subject_marks[all_subject_marks.index('410501')+6][:3]
        std_details_1['OR18'] = all_subject_marks[all_subject_marks.index('410501')+7][:3]
        std_details_1['TUT18'] = all_subject_marks[all_subject_marks.index('410501')+8][:3]
        std_details_1['TOT%18'] = all_subject_marks[all_subject_marks.index('410501')+9][:3]
        std_details_1['CRD18'] = all_subject_marks[all_subject_marks.index('410501')+10][:3]
        std_details_1['GRD18'] = all_subject_marks[all_subject_marks.index('410501')+11][:3]
        std_details_1['CP18'] = all_subject_marks[all_subject_marks.index('410501')+12][:3]
        std_details_1['GP18'] = all_subject_marks[all_subject_marks.index('410501')+13][:3]
        std_details_1['P&R18'] = all_subject_marks[all_subject_marks.index('410501')+14][:3]
        std_details_1['ORD18'] = all_subject_marks[all_subject_marks.index('410501')+15][:3]
    
    #print(all_subject_marks[all_subject_marks.index('410503')+2][:3])

    #HON-A.I. FOR BIG DATA ANA. for std1
    if '410503' in all_subject_marks:
        std_details_1['Insem19'] = all_subject_marks[all_subject_marks.index('410503')+2][:3]
        std_details_1['Ensem19'] = all_subject_marks[all_subject_marks.index('410503')+3][:3]
        std_details_1['Total19'] = all_subject_marks[all_subject_marks.index('410503')+4][:3]
        std_details_1['PW19'] = all_subject_marks[all_subject_marks.index('410503')+5][:3]
        std_details_1['PR19'] = all_subject_marks[all_subject_marks.index('410503')+6][:3]
        std_details_1['OR19'] = all_subject_marks[all_subject_marks.index('410503')+7][:3]
        std_details_1['TUT19'] = all_subject_marks[all_subject_marks.index('410503')+8][:3]
        std_details_1['TOT%19'] = all_subject_marks[all_subject_marks.index('410503')+9][:3]
        std_details_1['CRD19'] = all_subject_marks[all_subject_marks.index('410503')+10][:3]
        std_details_1['GRD19'] = all_subject_marks[all_subject_marks.index('410503')+11][:3]
        std_details_1['CP19'] = all_subject_marks[all_subject_marks.index('410503')+12][:3]
        std_details_1['GP19'] = all_subject_marks[all_subject_marks.index('410503')+13][:3]
        std_details_1['P&R19'] = all_subject_marks[all_subject_marks.index('410503')+14][:3]
        std_details_1['ORD19'] = all_subject_marks[all_subject_marks.index('410503')+15][:3]
        all_subject_marks.remove("410503")
        std_details_1['Insem20'] = all_subject_marks[all_subject_marks.index('410503')+2][:3]
        std_details_1['Ensem20'] = all_subject_marks[all_subject_marks.index('410503')+3][:3]
        std_details_1['Total20'] = all_subject_marks[all_subject_marks.index('410503')+4][:3]
        std_details_1['PW20'] = all_subject_marks[all_subject_marks.index('410503')+5][:3]
        std_details_1['PR20'] = all_subject_marks[all_subject_marks.index('410503')+6][:3]
        std_details_1['OR20'] = all_subject_marks[all_subject_marks.index('410503')+7][:3]
        std_details_1['TUT20'] = all_subject_marks[all_subject_marks.index('410503')+8][:3]
        std_details_1['TOT%20'] = all_subject_marks[all_subject_marks.index('410503')+9][:3]
        std_details_1['CRD20'] = all_subject_marks[all_subject_marks.index('410503')+10][:3]
        std_details_1['GRD20'] = all_subject_marks[all_subject_marks.index('410503')+11][:3]
        std_details_1['CP20'] = all_subject_marks[all_subject_marks.index('410503')+12][:3]
        std_details_1['GP20'] = all_subject_marks[all_subject_marks.index('410503')+13][:3]
        std_details_1['P&R20'] = all_subject_marks[all_subject_marks.index('410503')+14][:3]
        std_details_1['ORD20'] = all_subject_marks[all_subject_marks.index('410503')+15][:3]

    
    #HON-MACH. LEARN.& DATA SCI for std2
    if '410501' in all_subject_marks_2:
        std_details_2['Insem17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+2][:3]
        std_details_2['Ensem17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+3][:3]
        std_details_2['Total17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+4][:3]
        std_details_2['PW17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+5][:3]
        std_details_2['PR17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+6][:3]
        std_details_2['OR17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+7][:3]
        std_details_2['TUT17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+8][:3]
        std_details_2['TOT%17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+9][:3]
        std_details_2['CRD17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+10][:3]
        std_details_2['GRD17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+11][:3]
        std_details_2['CP17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+12][:3]
        std_details_2['GP17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+13][:3]
        std_details_2['P&R17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+14][:3]
        std_details_2['ORD17'] = all_subject_marks_2[all_subject_marks_2.index('410501')+15][:3]
        all_subject_marks_2.remove("410501")
        std_details_2['Insem18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+2][:3]
        std_details_2['Ensem18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+3][:3]
        std_details_2['Total18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+4][:3]
        std_details_2['PW18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+5][:3]
        std_details_2['PR18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+6][:3]
        std_details_2['OR18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+7][:3]
        std_details_2['TUT18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+8][:3]
        std_details_2['TOT%18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+9][:3]
        std_details_2['CRD18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+10][:3]
        std_details_2['GRD18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+11][:3]
        std_details_2['CP18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+12][:3]
        std_details_2['GP18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+13][:3]
        std_details_2['P&R18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+14][:3]
        std_details_2['ORD18'] = all_subject_marks_2[all_subject_marks_2.index('410501')+15][:3]
    
    #HON-A.I. FOR BIG DATA ANA. for std2
    if '410503' in all_subject_marks_2:
        std_details_2['Insem19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+2][:3]
        std_details_2['Ensem19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+3][:3]
        std_details_2['Total19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+4][:3]
        std_details_2['PW19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+5][:3]
        std_details_2['PR19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+6][:3]
        std_details_2['OR19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+7][:3]
        std_details_2['TUT19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+8][:3]
        std_details_2['TOT%19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+9][:3]
        std_details_2['CRD19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+10][:3]
        std_details_2['GRD19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+11][:3]
        std_details_2['CP19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+12][:3]
        std_details_2['GP19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+13][:3]
        std_details_2['P&R19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+14][:3]
        std_details_2['ORD19'] = all_subject_marks_2[all_subject_marks_2.index('410503')+15][:3]
        all_subject_marks_2.remove("410503")
        std_details_2['Insem20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+2][:3]
        std_details_2['Ensem20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+3][:3]
        std_details_2['Total20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+4][:3]
        std_details_2['PW20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+5][:3]
        std_details_2['PR20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+6][:3]
        std_details_2['OR20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+7][:3]
        std_details_2['TUT20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+8][:3]
        std_details_2['TOT%20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+9][:3]
        std_details_2['CRD20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+10][:3]
        std_details_2['GRD20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+11][:3]
        std_details_2['CP20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+12][:3]
        std_details_2['GP20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+13][:3]
        std_details_2['P&R20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+14][:3]
        std_details_2['ORD20'] = all_subject_marks_2[all_subject_marks_2.index('410503')+15][:3]
    
    #print(std_details_1.values())

    
    


    with open("Student_Details.csv","a",newline='') as csvfile:
        writer=csv.writer(csvfile)
        writer.writerow(std_details_1.values())
        if page!=end:
            writer.writerow(std_details_2.values())
        

    

    #print('410503' in all_subject_marks_2)
    #print(student_info_1)   #First half of the page  std details1
    #print(student_info_2)   #Second half of the page std details2
    #print(a1.split(':'))         #First half of the page  std scores1 CGPAs
    #print(a2,b2,c2)         #Second half of the page std scores2 CGPAs

    #print(all_subject_marks_2)
    #print(all_subject_marks_2[-51:])
