import textwrap
import pandas as pd
import pywhatkit
import datetime
import sys, os

# user file input with input validation
while True:
    file_input = input(f"name of student fail namelist file with .xlsx: ")
    if os.path.exists(file_input):
        df_input = pd.read_excel(file_input)
        break
    else:
        print(f"Invalid file. Please input a valid file name that ends with .xlsx.")  # 'Book1.xlsx'
while True:
    file_info = input(f"name of database file with .xlsx: ")
    if os.path.exists(file_info):
        df_info = pd.read_excel(file_info)
        break
    else:
        print(f"Invalid file. Please input a valid file name that ends with .xlsx.")  # 'Book1.xlsx'#pd.read_excel("student_database_sep.xlsx")
df_output = 'final_namelist.xlsx'
df_input.columns = df_input.columns.str.strip()

def send_messages():
    # visualisation of columns

    print(f"\nstudent fail namelist columns:")
    print(df_input.columns)

    # user column input with input validation
    while True:
        column_rename = input(f"name of column with student's name: ")
        if column_rename in df_input.columns:
            # calculate estimated time for messaging task
            print("length of specified column: " + str(df_input[column_rename].count())) #the time calculation must be before the use input column is renamed
            estimated_time = int(df_input[column_rename].count()) * 12
            print('estimated time to complete task in hr:min:sec : ' + str(datetime.timedelta(seconds=estimated_time)))
            df_input.rename(columns={column_rename: 'Student English Name'}, inplace=True) # 'F3 failed live quiz'
            break
        else:
            print(f"Invalid column name. Please input column name in student fail namelist.")

    # calculate estimated time for messaging task

    # df_input.rename(columns={column_rename: 'Student_English_Name'}, inplace=True)

    # creating a new column in df_output based on df_input and df_info overlap of 'Student_English_Name'
    df_output = pd.merge(df_input, df_info[['Student English Name', 'Emergency Contact No']], on='Student English Name',
                         how='left')

    df_output.dropna(subset='Emergency Contact No', inplace=True)

    df_output.to_excel('final_namelist.xlsx', index=False)

    # changing df_output['Emergency_Contact_No'] to +6 without '-' format, and removing numbers after "/"
    df_output['Emergency Contact No'] = '+6' + df_output['Emergency Contact No'].astype(str)

    df_output['Emergency Contact No'] = df_output['Emergency Contact No'].str.replace("-", "")

    df_output['Emergency Contact No'] = df_output['Emergency Contact No'].str.split('/').str[0]

    #reset column to original name to prevent redundancy in repeated usage
    df_input.rename(columns={"Student English Name": column_rename}, inplace=True)
    #print(df_input.columns)

    # random testing codes
    # print(df_output.loc [df_output['Emergency_Contact_No'].str.contains ("-")])
    # print(df_output['Emergency_Contact_No'])

    # for index in range(0, 15):
    #     print(df_output.loc [index, "Emergency_Contact_No"])

    # user input for whether failed message or correction message is sent
    while True:
        subject_choice = input(f"What subject this is for? \nscience(1), \nphysics(2) or \nchemistry(3): ")
        if subject_choice == "science" or subject_choice == "1":
            subject_name = "science "
        elif subject_choice == "physics" or subject_choice == "2":
            subject_name = "physics "
        elif subject_choice == "chemistry" or subject_choice == "3":
            subject_name = "chemistry "


        choice = input(f"\nfailed(1), \ndidn't pass up hwk chapter(2) or \ndidn't pass up correction(3): ")
        if choice == "failed" or choice == '1':
            test_name = input("name of test failed: ")
            test_con = "FAILED"
        elif choice == 'hwk' or choice == '2':
            test_name = input("name of absent hwk submission: ")
            test_con = "HWK CHAPTER SUBMISSION"
            chapter_number = input("chapter number= ")
        elif choice == '3':
            test_name = input("name of absent correction: ")
            test_con = "CORRECTION"

        # confirmation for execution
        confirmation = input(f"You are sending {subject_name}{test_con} for {column_rename}, {file_input} \n Proceed?(yes/no): ")
        if confirmation == "yes":
        # if function to send 3 types of personalised message, according to user input
            for index, row in df_output.iterrows():
                student_name = row[f'Student English Name']
                if choice == "failed" or choice == '1':
                    test = test_name
                    subject = subject_name
                    tutor_name = input("your name: ")
                    sent_successfully = True
                    message = textwrap.dedent(f'''
                    Hello {student_name}'s parent, this is {tutor_name}, tutor of cs tuition centre. \n
                    Your child has failed {subject}{test}. However, they did not hand in correction for it or attend tutorial. \n
                    tutorial time:
                    Thurs or Sun
                    F3: 7.30pm-8.30pm
                    F2: 8.30pm-9.30pm
                    F1: 9.30pm-10.30pm
                    link will be sent in announcement \n
                    Your child is required to hand in the correction to this WhatsApp chat by coping the answer and questions. File for correction is sent in CS APP science group.
                    https://chat.cstuition.com.my/home \n
            
                    Your attention is much appreciated. \U0001F60A ''')
                    # pywhatkit sending whatsapp messages
                    hp_number = row['Emergency Contact No']
                    try:
                        pywhatkit.sendwhatmsg_instantly(hp_number, message, wait_time=10, tab_close=True, close_time=2)
                        print(f"Message sent to {hp_number}")
                        sent_successfully = True
                    except Exception as e:
                        print(f"Failed to send message to {hp_number}: {str(e)}")
                        sent_successfully = False

                elif choice == 'hwk submission'or choice == '2':
                    test= test_name
                    subject = subject_name
                    tutor_name = input("your name: ")
                    sent_successfully = True
                    message = textwrap.dedent(f'''
                                Hello {student_name}'s parent, this is {tutor_name}, tutor of cs tuition centre. \n
                                Your child did not hand in {subject}{test}.
        
                                Your child is required to take picture of chapter {chapter_number} exercise book subjective questions and hand in to this WhatsApp chat. \n
        
                                Your attention is much appreciated. \U0001F60A''')
                    # pywhatkit sending whatsapp messages
                    hp_number = row['Emergency Contact No']
                    try:
                        pywhatkit.sendwhatmsg_instantly(hp_number, message, wait_time=10, tab_close=True, close_time=2)
                        print(f"Message sent to {hp_number}")
                        sent_successfully = True
                    except Exception as e:
                        print(f"Failed to send message to {hp_number}: {str(e)}")
                        sent_successfully = False

                elif choice == '3' :
                    test = test_name
                    subject = subject_name
                    sent_successfully = True
                    tutor_name = input("your name: ")
                    message = textwrap.dedent(f'''
                    Hello {student_name}'s parent, this is {tutor_name}, tutor of cs tuition centre. \n
                    Your child did not hand in {subject}{test}.
            
                    Your child is required to hand in correction to this WhatsApp chat by coping answer and questions. File for correction is sent in CS APP science group.
                    https://chat.cstuition.com.my/home \n
            
                    Your attention is much appreciated. \U0001F60A''')
                    # pywhatkit sending whatsapp messages
                    hp_number = row['Emergency Contact No']
                    try:
                        pywhatkit.sendwhatmsg_instantly(hp_number, message, wait_time=10, tab_close=True, close_time=2)
                        print(f"Message sent to {hp_number}")
                        sent_successfully = True
                    except Exception as e:
                        print(f"Failed to send message to {hp_number}: {str(e)}")
                        sent_successfully = False

            if sent_successfully == True:
                break
        elif confirmation == 'no':
            continue

while True:
    send_messages()
    repeat_command = input("do you want to send more messages using this input file?(yes/no): ")
    if repeat_command == 'no':
        break

input("Press Enter to exit...")