#-*-coding:utf-8-*-
#__author__='maxiaohui'

'''
As in resource files, text use variable which value is saved in Data Source(Excel)
The script is used to build available resource file
'''

import codecs
import time,os,shutil,re,sys,csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

reload(sys)
sys.setdefaultencoding( "utf-8" )

#配置文件  1、产品提供的数据源 2、开发提供的template 3、结果反馈文件
#localPath=os.path.abspath('.')   #当前文件的上一级目录
localPath=os.path.dirname(__file__)

#源文件目录
csv_source= os.path.join(localPath,"Bbox_fontDB.csv")
language_column={"cn":3,"en":4,"japan":5}
#反馈文件
csv_config=os.path.join(localPath,'config.csv')
    #print csv_config
#android项目目录
projectPath=os.path.join(localPath,"packages")


#邮件信息配置
sender='maxiaohui@beeboxes.com'
to_receiver=['maxiaohui@beeboxes.com']
to_receiver=['xuliyang@beeboxes.com']
cc_reciver=['maxiaohui@beeboxes.com']
receiver = to_receiver + cc_reciver
email_text = '''
        Hi all
        ----------本邮件由Python2.7脚本自动发送----------
        附件中是资源文件build时缺少的字段，请尽快补齐。

        备注：以下是自动生成的资源文件：
        '''

class easyCSV():
    '''
    csv基本操作，读，写，搜索
    '''

    # 打开文件或者新建文件（如果不存在的话）
    def __init__(self, filename=None):
        self.filename = filename
        if not filename:   #创建新文件
            csvfilename='Resource_notFound_%s.csv' % time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
            self.filename=os.path.join(localPath,csvfilename)
            self.fopen = open(self.filename,"wb")
            self.fopen.write(codecs.BOM_UTF8)
            self.csvHandle = csv.writer(self.fopen,delimiter='^')
            csv_head = ["文件模板名","行数","问题行","备注","语言"]
            self.csvHandle.writerow(csv_head)
        else:    #读取老文件
            with open(filename,"rb") as cf:
                self.csvHandle = cf.readlines()

    def searchTextById(self,id,language):
        status=False
        for line in self.csvHandle:
            line=line.encode('utf-8')
            line=line.split('^')
            if line[0]==id:
                text=line[language_column[language]-1]
                status=True
                break
        if not status:
            text=id
        # if id==u"T00440":
        #     print(text)
        #     print(type(text))
        return text

class fileHandler():
    '''
    文件的copy，复制，删除，移动等
    '''
    def copy_file(self,source,aimpath,postfix=""):  #把源文件copy到目标目录下，文件名不变或加后缀
        path,filename=os.path.split(source)
        try:
            filename=filename.split(".")[0]+postfix+'.'+filename.split(".")[1]
        except IndexError:   #处理没有后缀名的配置文件
            pass
        copyed_file = os.path.join(aimpath, filename)
        shutil.copy(source,os.path.join(aimpath,filename))
        return copyed_file

    def move_file(self,source,aim):
        shutil.move(source,aim)

    def del_file(self,source):
        os.remove(source)

    def del_folder(self,folderPath): #删除空文件夹
        os.rmdir(folderPath)

    def make_dir(self,path):  #创建一个文件夹
        os.mkdir(path)

class resourceBuilder():  #一个类对应一个template file，多种语言
    '''
    支持替换文件中的变量T000123，生成目标resource文件,自动反馈问题结果
    '''
    fHandler=fileHandler()

    #以下是初始化部分，准备各种文件
    def __init__(self,configCSV,sourceCSV):  #传入参数：配置文件xml，源文件（xuliyang提供)
        self.config=configCSV
        self.source=sourceCSV
        self.result_folder = self.prepare_result_folder()
        self.prepare_source_feedback_csv()
        self.config_list = self.get_config()

    def prepare_source_feedback_csv(self): #准备配置文件，源文件，feedback文件
        self.csv_config = easyCSV(self.config)   #配置文件，只用来读取。 增加语言或者增加模板直接新加一行
        self.csv_source = easyCSV(self.source)            #源文件，只用来读取
        self.csv_feedback = easyCSV()  #反馈给xuliyang的文件，只用来读写

    def get_config(self):   #从配置文件中获取配置列表
        config_list=[]
        self.csv_config.csvHandle.pop(0)  #删除标头
        for row in self.csv_config.csvHandle:
            row=row.strip() #移除前后的空字符 \r\n
            d = dict(zip(['template_path', 'language', 'aim_path', 'app_name'], row.split(",")))
            config_list.append(d)
        return config_list

    def prepare_result_folder(self):  #生成结果文件夹
        result_folder = os.path.join(localPath, "temporary_folder")
        try:
            os.mkdir(result_folder)
        except OSError:
            shutil.rmtree(result_folder)
            os.mkdir(result_folder)
        return result_folder

    #以下部分为主流程，获取待处理文件，根据变量生成目标文件，放到制定目录中
    def get_files(self,template_file,result_folder): #template_path, aim_path来自于配置文件  路径完整
        template_file=projectPath+template_file
        template_file = self.fHandler.copy_file(template_file,result_folder)
        return template_file    #处理过程中的文件

    def getValue_from_csv(self, id, language):  # 根据id：T00054，获取value值，如果找不到，返回id
        return self.csv_source.searchTextById(id,language)

    def write_line_feedback(self,feedbackList):  #对于有问题的条目，写入内容
        self.csv_feedback.csvHandle.writerow(feedbackList)

    def get_values(self,template_file,language):   #处理两种异常：有变量号找不到值，新增变量{{测试文本}}
        file_data = ""
        with codecs.open(template_file,"r","utf-8") as f:
            line_no=1
            for line in f:
                txtVa = re.findall(r"(T\d{5,10})", line)   #注意！变量标号的数字位数最短不能少于5位数字
                txtNewCreated = re.findall(r"{{(.*?)}}", line)
                if txtVa:   #在文件中找到变量 T00234
                    old_str = txtVa[0]
                    new_str = self.getValue_from_csv(old_str, language)
                    if old_str == new_str:  #代表excel中没有找到相关变量，记录到反馈文件中
                        feedbackList=[self.templateCurrentRun,line_no,line, "未找到条目:%s"%new_str,language]
                        self.write_line_feedback(feedbackList)
                    if new_str=="":   #代表excel中有该变量名，但是字符串为空，例如英文没有翻译
                        feedbackList=[self.templateCurrentRun, line_no, line, "%s未翻译:%s" %(language,old_str), language]
                        self.write_line_feedback(feedbackList)
                        new_str=old_str
                    if new_str=="null!":   #针对英文有些太长不需要显示的用跟着字符表示空
                        new_str=""
                    line = line.replace(old_str, new_str)
                if txtNewCreated: #是新添加的变量，中文替换，其他语言置空
                    line = line.replace("{{", "")
                    line = line.replace("}}", "")
                    feedbackList=[self.templateCurrentRun, line_no, line, "新增条目:%s" % txtNewCreated[0], language]
                    self.write_line_feedback(feedbackList)
                line_no += 1
                file_data += line

        with codecs.open(template_file, "w", "utf-8") as f:
            f.write(file_data)

    def clear_down(self,template_file,aim_file): #move template文件到指定目录，保存feedback文件
        self.fHandler.move_file(template_file,projectPath+aim_file)
        global email_text
        email_text=email_text+"\n"+aim_file
        print("Generate Target File：%s"%(projectPath+aim_file))

    def send_email(self,sender, receiver,text,attachment):
        # 第三方 SMTP 服务
        mail_host = "mail.beeboxes.com"  # 设置服务器
        mail_user = "maxiaohui"  # 用户名
        mail_pass = "Chinaxin1234"  # 口令

        content = MIMEText(text,"plain","utf-8")
        message = MIMEMultipart()
        message.attach(content)
        message['From'] = sender
        message['To'] = ";".join(to_receiver)
        message['Cc'] = ";".join(cc_reciver)
        subject = '资源文件反馈_%s' % time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
        message['Subject'] = Header(subject, 'utf-8')
        xlsx = MIMEApplication(open(attachment, 'r').read())
        xlsx["Content-Type"] = 'application/octet-stream'
        xlsx.add_header('Content-Disposition', 'attachment', filename=os.path.split(attachment)[1])
        message.attach(xlsx)
        try:
            smtpObj = smtplib.SMTP()
            smtpObj.connect(mail_host, 25)  # 25 为 SMTP 端口号
            smtpObj.login(mail_user, mail_pass)
            smtpObj.sendmail(sender, receiver, message.as_string())
            print "Send email successfuly"
        except  Exception as e:
            print(e)

    def tear_down(self,sendEmail=True):  #保存feedback文件，发送邮件，删除feedback文件，清空临时文件夹
        self.csv_feedback.fopen.close()
        if sendEmail:
            self.send_email(sender, receiver, email_text,self.csv_feedback.filename)
        self.fHandler.del_file(self.csv_feedback.filename)
        self.fHandler.del_folder(self.result_folder)

    def build_all_config(self): #把excel中所有的配置条目，从第二行开始，都build了
        for i in self.config_list:
            try:
                self.templateCurrentRun=i['template_path']
                template_file = self.get_files(i['template_path'], rd.result_folder)
                self.get_values(template_file, i["language"])
                self.clear_down(template_file, i['aim_path'])
            except IOError:
                print("无模板文件：%s，不处理" % i['template_path'])
        self.tear_down()

    def build_config_column(self,column_no): #把excel某个条目，从第二行开始，都build了
        column=1
        for i in self.config_list:
            column+=1
            if column==column_no:
                try:
                    self.templateCurrentRun = i['template_path']
                    template_file = self.get_files(i['template_path'], rd.result_folder)
                    self.get_values(template_file, i["language"])
                    self.clear_down(template_file, i['aim_path'])
                except IOError:
                    print("无模板文件：%s，不处理" % i['template_path'])
        self.tear_down(sendEmail=False)

    def build_config_app(self,app_name): #针对app进行build语言文件
        for i in self.config_list:
            if i["app_name"]==app_name:
                try:
                    self.templateCurrentRun = i['template_path']
                    template_file = self.get_files(i['template_path'], rd.result_folder)
                    self.get_values(template_file, i["language"])
                    self.clear_down(template_file, i['aim_path'])
                except IOError:
                    print("无模板文件：%s，不处理" % i['template_path'])
        self.tear_down(sendEmail=False)

if __name__ == "__main__":
    rd=resourceBuilder(csv_config,csv_source)   #excel_config由开发配置， excel_source由xuliyang提供
    try:
        if sys.argv[1]:
            rd.build_config_app(sys.argv[1])       #带了app参数的话，会只build当前app
    except IndexError:
        rd.build_all_config()
        # rd.build_config_column(5)                   #如果只想生成一个文件，可以用这条语句
