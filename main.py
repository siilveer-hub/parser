import base64
import email
import imaplib
import os
import re
import shutil

from email.header import decode_header
from docx2pdf import convert

def stud_list_in_group():
    # cwd = os.getcwd()
    file_list = os.listdir(path=".")

    for file in file_list:
        if "txt" in file:
            group = file.split('.')[0]
            with open(file, 'r',encoding ='utf-8') as f:
                stud = [line.strip() for line in f]
                f.close()
    return stud, group

def moving_groups(filename):
    Group_list =[]
    with open(filename, 'r', encoding='UTF-8') as file:
        while line := file.readline():
            tmp=line.rstrip().split()
            tmp = tmp[0]+'_'+tmp[1]+tmp[2]
            Group_list.append(tmp)

    group_dir = filename[:-4]

    dest_path = os.path.join(os.getcwd(),group_dir)
    cwd = os.getcwd()

    base_path = os.path.join(cwd,'attachment')
    dir_list = os.listdir(base_path)
    
    if os.path.exists(dest_path):
        pass    
    else:
        os.mkdir(dest_path)
        print(f"Директория {dest_path} создана")

    for dir in dir_list:
        if dir in Group_list:
            source = os.path.join(base_path, dir)
            path = os.path.join(dest_path,dir)
            get_files = os.listdir(source)     
            if os.path.exists(path):
                pass    
            else:
                os.mkdir(path)
                print(f"Директория {path} создана")  
            file_list = os.listdir(source)
            for f in get_files:
                old = os.path.join(base_path, dir,f)
                new = os.path.join(path, f)
                if os.path.exists(new):
                    os.remove(new)
                shutil.move(old, path)
            shutil.rmtree(source)        
    print("Файлы студентов группы ", group_dir, "перемещены.")
    return
  
class student:
    def __init__(self,sender, subject):        
        self.sender = sender
        self.subject = subject

    def extract_email(self):
        """Функция декодирования почты и ФИО отправителя"""
        text = self.sender
        self.sender_mail = text.split(' ')[-1][1:-1]
        return self.sender_mail

    def extract_name(self):
        """Функция декодирования почты и ФИО отправителя"""
        text = self.sender
        self.name = decode_header(text)[0][0].decode('utf-8')
        return self.name
    
    def extract_subject(self):
        text = self.subject
        """Функция декодирования заголовка письма"""
        if text ==  None:
            self.subj = 'Без темы'
            return self.subj
        else:
            code = decode_header(message["Subject"])[0][-1]
            if code == 'utf-8':
                self.subj = decode_header(text)[0][0].decode('utf-8')
                return self.subj
            else:
                self.subj = decode_header(text)[0][0]
                return self.subj
            
    def extract_filename(self, message):
        """Функция поиска почтовых вложений и вывод их названий"""
        files_list = []
        for part in message.walk():
            file_name = part.get_filename()
            if file_name == None:
                continue
            else:
                if decode_header(file_name)[0][1] == None:
                    file_name = decode_header(file_name)[0][0]
                    files_list.append(file_name)                    
                else:
                    file_name = decode_header(file_name)[0][0].decode('utf-8')
                    files_list.append(file_name)
                self.download_attach(part,file_name)    
                self.create_studir(self.name)

                ## Проверь эту часть кода и разберись!
                # path = "attachment"
                # if bool(file_name):
                #     filePath = os.path.join(path, file_name)
                #     if not os.path.isfile(filePath) :
                #         fp = open(filePath, 'wb')
                #         fp.write(part.get_payload(decode=True))
                #         fp.close()
                # self.download_attach(part,file_name)
        return files_list

    def download_attach(self, part, file_name):
        """Функция загрузки почтовых вложений"""
        path = "attachment"
        if bool(file_name):
            filePath = os.path.join(path, file_name)
            if not os.path.isfile(filePath) :
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()

    def create_studir(self,stud_name):
        """Функция создания директории студента и перемещение файлов в директорию"""
        stud_dir = stud_name.split()[0]+'_'+stud_name.split()[1][0]+stud_name.split()[2][0]
        cwd = os.getcwd()
        dest_path = os.path.join(cwd,'attachment',stud_dir)
        if os.path.exists(dest_path):
            pass
        else:
            os.mkdir(dest_path)
            print(f"Директория {dest_path} создана")

        attach_dir = os.path.join(cwd,'attachment')
        os.chdir(attach_dir)
        file_list = os.listdir(path=".")

        for file in file_list:
            if "docx" in file:
                convert(file,stud_dir)
                os.remove(file)
            if "pdf" in file:
                src = os.path.join(attach_dir,file)
                dest = os.path.join(dest_path,file)
                try:
                    os.replace(src,dest)
                except:
                    os.remove(src,dest)
            if "png" in file:
                src = os.path.join(attach_dir,file)
                dest = os.path.join(dest_path,file)
                try:
                    os.replace(src,dest)
                except:
                    pass
            if "xmcd" in file:
                src = os.path.join(attach_dir,file)
                dest = os.path.join(dest_path,file)
                try:
                    os.replace(src,dest)
                except:
                    pass

        os.chdir(cwd)
        
#########################################################################################


imap_server = "***"
email_address = "***"
password = "***"

mail= imaplib.IMAP4_SSL(imap_server)
mail.login(email_address,password)

mail.select("Inbox")
# stud_list, group = stud_list_in_group()
_, msgnums = mail.search(None,'(SINCE "31-Aug-2022" BEFORE "31-Dec-2022")')
print(msgnums)


# email_num = [946, 966, 996, 1001, 1009, 1023]
# email_num = [5, 18, 20, 22, 996, 1009]
# email_num = [5, 231]
# email_num = [5]

# for i in email_num:
#     idx = msgnums[0].split()
#     _, data = mail.fetch(idx[i], '(RFC822)')
#     message = email.message_from_bytes(data[0][1])

#     stud = student(message["From"],message["Subject"])
#     stud_mail = stud.extract_email()

#     if "student" in stud_mail:
#         stud_name =  stud.extract_name()        
#         title = stud.extract_subject()
#         file_names = stud.extract_filename(message)     

#         print(f"")
#         print(f"Номер сообщения: {idx[i]}")  
#         print(f"Дата получения: {message['Date']}")
#         print(f'ФИО: {stud_name}')
#         print(f'Почта отправителя: {stud_mail}')
#         print(f'Тема письма: {title}')
#         print(f'Список вложений: {file_names}')





#########################################################################################
# Основная часть кода
# stud_list, group = stud_list_in_group()

for msgnum in msgnums[0].split():
    _, data = mail.fetch(msgnum, '(RFC822)')    
    message = email.message_from_bytes(data[0][1])
    stud = student(message["From"],message["Subject"]) 
    stud_mail = stud.extract_email()

    if "student" in stud_mail:
        stud_name =  stud.extract_name()
        title = stud.extract_subject()
        file_names = stud.extract_filename(message)       

        print(f"")
        print(f"Номер сообщения: {msgnum}")  
        print(f"Дата получения: {message['Date']}")
        print(f'ФИО: {stud_name}')
        print(f'Почта отправителя: {stud_mail}')
        print(f'Тема письма: {title}')
        print(f'Список вложений: {file_names}')
# filename ='СМ5-41.txt'
# moving_groups(filename)


    
