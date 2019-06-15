# coding:utf-8
# Author:Fenn
# blog: fennbk.com
from PyQt5 import QtCore,QtGui,QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import pyqtSignal,Qt
import qtawesome,sys,requests
from PyQt5.QtGui import QMovie,QFont
from PyQt5.QtGui import  QPixmap,QIcon,QCursor
from lxml import etree
class QHead_Button(QPushButton):
    def __init__(self, *args):
        super(QHead_Button, self).__init__(*args)
        self.setFont(QFont("Webdings"))
        self.setFixedWidth(40)
class Show_Imageis(QWidget):
    Signal = pyqtSignal(str)
    def __init__(self):
        super(Show_Imageis,self).__init__()
        self._MoveVarl = False
    def wheelEvent(self, event):
        if (event.angleDelta()).y() > 0:
            self.Signal.emit('up')
        else:
            self.Signal.emit('down')

    def mousePressEvent(self, event):
        self._MoveVarl = True
        self.move_DragPosition = event.globalPos() - self.pos()
        event.accept()
    def mouseMoveEvent(self, QMouseEvent):
        if self._MoveVarl:
            self.move(QMouseEvent.globalPos() - self.move_DragPosition)
            QMouseEvent.accept()
    def mouseReleaseEvent(self, QMouseEvent):
        self._MoveVarl = False
# class Load_window(QWidget):
#     def __init__(self,*args,**kwargs):
#         super(Load_window,self).__init__(*args,**kwargs)
#         self.load_1 = QLineEdit(self)
#         self.load_1.resize(940, 680)
#         self.load_1.move(10, 10)
#         self.load_1.setEnabled(False)
#         self.load_1.setStyleSheet('''
#                     background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:0, y2:1, stop:0.504386 rgba(131, 131, 131, 174));
#                              border-top-right-radius:10px;
#                              border-bottom-right-radius:10px;
#                      ''')
#         self.load_2 = QLabel(self)
#         self.load_2.setGeometry(470, 340, 100, 100)
#         gif = QMovie('C:\\Users\\Administrator\\Desktop\\load.gif')
#         self.load_2.setMovie(gif)
#         gif.start()
#         left_cont = 0
#         cont = 0
#         for item in '正在加载网络资源请稍后....':
#             self.loading_text.append(QLabel(item, self))
#             self.loading_text[cont].move(400 + left_cont, 340)
#             self.loading_text[cont].setObjectName('QlabelLoad')
#             cont += 1
#             left_cont += 20
#         self.setStyleSheet('''
#                          QLabel#QlabelLoad{
#                          font-size: 16px;
#                          font-weight:600;
#                          color:#FFFFFF
#                          }
#
#                      ''')

class PPT_App(QtWidgets.QMainWindow):
    def __init__(self):
        super(PPT_App,self).__init__()
        self.url = 'http://www.1ppt.com'
        self._MoveVarl = False
        self.resize(1100,700)
        self.gridLayout_infor = QtWidgets.QGridLayout(self)
        self.gridLayout_load = QtWidgets.QGridLayout(self)
        self.main_widget = QtWidgets.QWidget()  # 创建窗口主部件
        self.main_layout = QtWidgets.QGridLayout()  # 创建主部件的网格布局
        self.main_widget.setLayout(self.main_layout)  # 设置窗口主部件布局为网格布局

        self.left_widget = QtWidgets.QWidget()  # 创建左侧部件
        self.left_widget.setObjectName('left_widget')
        self.left_layout = QtWidgets.QGridLayout()  # 创建左侧部件的网格布局层
        self.left_widget.setLayout(self.left_layout) # 设置左侧部件布局为网格

        self.right_widget = QtWidgets.QWidget() # 创建右侧部件
        self.right_widget.setObjectName('right_widget')
        self.right_layout = QtWidgets.QGridLayout()
        self.right_widget.setLayout(self.right_layout) # 设置右侧部件布局为网格

        self.main_layout.addWidget(self.left_widget,0,0,12,2) # 左侧部件在第0行第0列，占8行3列
        self.main_layout.addWidget(self.right_widget,0,2,12,10) # 右侧部件在第0行第3列，占8行9列
        self.setCentralWidget(self.main_widget) # 设置窗口主部件
        self.left_visit = QtWidgets.QPushButton(qtawesome.icon('fa.xing', color='red'),'免费PPT下载\n工具 V1.0')  # 空白按钮
        self.left_label_1 = QtWidgets.QPushButton("模板分类")
        self.left_label_1.setObjectName('left_label')
        self.Men_list = self.Request_info('Get_Men')
        self._ClosButton = QHead_Button(b'\xef\x81\xb2'.decode("utf-8"), self)
        self._ClosButton.setObjectName("ClosButton")
        self._ClosButton.setCursor(QCursor(Qt.PointingHandCursor))
        self._ClosButton.move(self.width() - self._ClosButton.width() - 12, 14)
        self._MinimumButton = QHead_Button(b'\xef\x80\xb0'.decode("utf-8"), self)
        self._MinimumButton.setObjectName("MinMaxButton")
        self._MinimumButton.setCursor(QCursor(Qt.PointingHandCursor))
        self._MinimumButton.move(self.width() - self._ClosButton.width() - 40, 14)
        self._InfoButton = QHead_Button(qtawesome.icon('fa.user', color='#FFFFFF'),'', self)
        self._InfoButton.setObjectName("MinMaxButton")
        self._InfoButton.setCursor(QCursor(Qt.PointingHandCursor))
        self._InfoButton.move(self.width() - self._ClosButton.width() - 70, 14)
        self._InfoButton.clicked.connect(self.AuthorInformation)
        self._MinimumButton.clicked.connect(self.showMinimized)
        self._ClosButton.clicked.connect(self.close)
        self.left_button_list=[]
        for item in self.Men_list:
            self.left_button_list.append(QtWidgets.QPushButton(qtawesome.icon('fa.reorder', color='red'),list(item.keys())[0]))
        self.left_layout.addWidget(self.left_visit, 0, 1, 1, 1)
        self.left_layout.addWidget(self.left_label_1, 1, 0, 1, 3)
        cont = 2
        for item in self.left_button_list:
            self.left_layout.addWidget(item, cont, 0, 1, 3)
            item.setObjectName('left_button')
            item.setCheckable(True)
            item.clicked.connect(self.button_Cilck)
            cont +=1
        self.right_recommend_label = QtWidgets.QLabel("分类列表")
        self.right_recommend_label.setObjectName('right_lable')
        self.right_layout.addWidget(self.right_recommend_label, 0, 0, 1, 9)
        self.left_visit.setStyleSheet(
            '''QPushButton{color:#F7D674;border-radius:5px;}''')
        self.left_widget.setStyleSheet('''
        QWidget#left_widget{
            background:#303133;
            border-top-left-radius:10px;
            border-bottom-left-radius:10px;
        }    
        QPushButton{border:none;color:#828282;font-weight:500;}
        QPushButton#left_label{
                border:none;
                border-bottom:1px solid white;
                font-size:18px;
                font-weight:700;
                font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
            }
            QPushButton#left_button:hover{border-left:4px solid #EEB422;font-weight:700;color:#EEB422}
        ''')
        self.right_widget.setStyleSheet('''
            QWidget#right_widget{
                color:#232C51;
                background-color: qlineargradient(spread:pad, x1:0.524597, y1:1, x2:0.524, y2:0, stop:0 rgba(255, 255, 255, 255), stop:0.0384615 rgba(216, 255, 249, 255), stop:0.387987 rgba(157, 248, 255, 255), stop:0.631494 rgba(156, 237, 255, 255), stop:1 rgba(112, 192, 255, 255));
                border-top:1px solid darkGray;
                border-bottom:1px solid darkGray;
                border-right:1px solid darkGray;
                border-top-right-radius:10px;
                border-bottom-right-radius:10px;
            }
            QLabel#right_lable{
                border:none;
                font-size:16px;
                font-weight:700;
                font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
            }
        ''')
        self.setStyleSheet('''
        QPushButton#ClosButton:hover{
            color:#FF6347
        }
        QPushButton#MinMaxButton:hover {
            color:#FF6347
        }
        QPushButton
{

    border-color: #31363b;
    border-radius:50%;
    color: #FFFFFF;


    border-width: 0px;

    border-style: solid;
    padding: 5px;
    border-radius: 2px;
    outline: none;
}
        ''')
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.main_layout.setSpacing(0)
        self.left_button_list[0].setChecked(True)
        self.Start_Boot= self.left_button_list[0].text()
        self.button_Cilck()
    def mousePressEvent(self, event):
        try:
            self.Show_Widget.close()
        except:
            pass
        if event.y() < 42:
            self._MoveVarl = True
            self.move_DragPosition = event.globalPos() - self.pos()
            event.accept()
    def mouseMoveEvent(self, QMouseEvent):
        if self._MoveVarl:
            self.move(QMouseEvent.globalPos() - self.move_DragPosition)
            QMouseEvent.accept()
    def mouseReleaseEvent(self, QMouseEvent):
        self._MoveVarl = False
    def Request_info(self,GET_TYPE,url_data='',Name=''):
        default_headers = {
            'Accept': 'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'charset': 'UTF-8',
            # 'Upgrade-Insecure-Requests':'1',用于让浏览器自动升级请求从http到https，一定不要添加，上传文件后无法发送文件
            'Accept-Language': 'en-US,en;q=0.8,zh-CN;q=0.5,zh;q=0.3',
            # 'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
            'Connection': 'keep-alive',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0',
        }
        self.Menu_page_kj = False
        self.Menu_page_sc = False
        try:
            if GET_TYPE == 'Get_Men':

                title_ur = (etree.HTML((requests.get(self.url,headers=default_headers)).content)).xpath('//*[@class="col_nav i_nav clearfix"]/ul')
                Men_list = []
                i = 0
                for tite_text in title_ur:
                    Men_list.append({tite_text.xpath('li/text()')[0]: []})
                    for key, url in zip(tite_text.xpath('li/a/text()'), tite_text.xpath('li/a/@href')):
                        Men_list[i][tite_text.xpath('li/text()')[0]].append({key: url})
                    i += 1
                return Men_list
            elif GET_TYPE == 'Get_Details':
                body_url = (etree.HTML((requests.get(self.url + url_data,headers=default_headers)).content)).xpath('/html/body/div[5]/dl/dd/ul/li')
                Men_list = []
                if not body_url:
                    body_url = (etree.HTML((requests.get(self.url + url_data,headers=default_headers)).content)).xpath(
                        '/html/body/div[4]/div[1]/dl/dd/ul/li')
                    if not body_url:
                        body_url = (
                            etree.HTML(((requests.get(self.url + url_data,headers=default_headers)).content).decode('gbk'))).xpath('/html/body/div[5]/div[1]/dl/dd/ul/li')
                        Men_list = []
                        for body_text in body_url:
                            Men_list.append({body_text.xpath('a/img/@alt')[0]: [body_text.xpath('a/img/@src')[0],
                                                                                body_text.xpath('a/@href')[0]]})
                        return Men_list
                    for body_text in body_url:
                        Men_list.append({body_text.xpath('a/img/@alt')[0]: [body_text.xpath('a/img/@src')[0],
                                                                            body_text.xpath('a/@href')[0]]})
                    return Men_list


                try:
                    for body_text in body_url:
                        Men_list.append({body_text.xpath('a/img/@alt')[0]: [body_text.xpath('a/img/@src')[0],body_text.xpath('a/@href')[0]]})
                    return Men_list
                except:
                    Men_list=[]
                    body_url = (etree.HTML((requests.get(self.url + url_data,headers=default_headers)).content)).xpath(
                        '/html/body/div[4]/div/ul/li/a')
                    for item in body_url:
                        Men_list.append({item.xpath('text()')[0]: item.xpath('@href')[0]})
                    return {'kejian': Men_list}
            elif GET_TYPE == 'Get_Images':
                    photo = QPixmap()
                    photo.loadFromData((requests.get(url_data,headers=default_headers)).content)
                    return photo
            elif GET_TYPE == 'Get_Download_url':
                return (etree.HTML((requests.get(self.url+url_data,headers=default_headers)).content)).xpath('/html/body/div[4]/div[1]/dl/dd/ul/li/a/@href')[0]
            elif GET_TYPE == 'Download_PPT':
                Download_Dir = QFileDialog.getExistingDirectory(self,
                                                               "选择下载路径",
                                                               "C:\\Users\\Administrator\\Desktop")
                res = requests.get(url_data,headers=default_headers)
                res.raise_for_status()
                playFile = open(Download_Dir+'\\'+Name+'.'+url_data.split('.')[-1], 'wb')
                for chunk in res.iter_content(100000):
                    playFile.write(chunk)
                playFile.close()
            elif GET_TYPE == 'show_ppt':
                images_list=[]
                Download_url = (etree.HTML((requests.get(self.url+url_data,headers=default_headers)).content)).xpath('/html/body/div[4]/div[1]/dl/dd/div[2]/p[1]/img')
                if not Download_url:Download_url = (etree.HTML((requests.get(self.url+url_data)).content)).xpath('/html/body/div[4]/div[1]/dl/dd/div[2]/img')
                for item in Download_url:
                    images_list.append(item.xpath('@src')[0])
                return images_list
            elif GET_TYPE == 'Get_page':
                body_url = (etree.HTML((requests.get(self.url + url_data,headers=default_headers)).content)).xpath(
                    '/html/body/div[5]/dl/dd/div/ul/li')
                if not body_url:
                    self.Menu_page_sc = True
                    body_url = (etree.HTML((requests.get(self.url + url_data,headers=default_headers)).content)).xpath(
                        '/html/body/div[4]/div[1]/dl/dd/div/ul/li')
                    if not body_url:
                        self.Menu_page_kj = True
                        body_url =(etree.HTML((requests.get(self.url + url_data,headers=default_headers)).content.decode('gbk'))).xpath(
                            '/html/body/div[5]/div[1]/dl/dd/div/ul/li')
                        try:
                            return body_url[-1].xpath('a/@href')[0].split('.')[0].split('_')[-1]
                        except:
                            return body_url[-2].xpath('text()')[0]
                    try:
                        return body_url[-1].xpath('a/@href')[0].split('/')[-2]
                    except:
                        return body_url[-2].xpath('text()')[0]
                try:
                    return body_url[-1].xpath('a/@href')[0].split('.')[0].split('_')[-1]
                except:
                    return body_url[-2].xpath('text()')[0]
        except Exception as E:
            print(E)
    def button_Cilck(self):
        self.right_recommend_label.setText("PPT分类")
        try:
            self.right_recommend_widget.setParent(None)
            self.tableWidget.setParent(None)
            self.pushButton_down.setParent(None)
            self.pushButton_up.setParent(None)
            self.Tpage_lable.setParent(None)
            self.right_layout.removeWidget(self.right_recommend_widget)
        except:
            pass
        self.right_recommend_widget = QtWidgets.QWidget()
        self.right_layout.addWidget(self.right_recommend_widget, 1, 0, 2, 9)
        self.right_recommend_layout = QtWidgets.QGridLayout()
        self.right_recommend_widget.setLayout(self.right_recommend_layout)

        self.right_recommend_widget.setStyleSheet(
            '''
                QToolButton{border:none;}
                QToolButton:hover{border-bottom:2px solid #F76677;}
            ''')
        body_data=[]

        for item in self.Men_list:
            if self.Start_Boot:
                try:
                    body_data = item[self.Start_Boot]
                    self.Start_Boot = False
                except:
                    pass
            else:
                try:
                    body_data = item[self.sender().text()]
                except:
                    pass
        self.recommend_button_list=[]
        self.Body_url=body_data
        cont = 0
        cont2 = 0
        for item in body_data:
            self.recommend_button_list.append(QtWidgets.QToolButton())
            nxet_cont= len(self.recommend_button_list)-1
            self.recommend_button_list[nxet_cont].setText(list(item.keys())[0])
            if cont <= 4:
                self.right_recommend_layout.addWidget(self.recommend_button_list[nxet_cont],cont2,cont)
                cont += 1
            else:
                cont2 += 1
                self.right_recommend_layout.addWidget(self.recommend_button_list[nxet_cont], cont2, 0)
                cont=1
        for item in self.recommend_button_list:
            item.clicked.connect(self.PPT_list)
    def Download_Button(self,name,id,Dtype=False):
        Widget = QWidget()
        if Dtype:
            kejian_button = QPushButton(qtawesome.icon('fa.book'), name, Widget)
            kejian_button.setCursor(QCursor(Qt.PointingHandCursor))
            kejian_button.clicked.connect(self.PPT_list)
            kejian_button.setStyleSheet('''
            color:black;font-weight:500;
            
            ''')
            hLayout = QHBoxLayout()
            hLayout.addWidget(kejian_button)
            hLayout.setContentsMargins(5, 2, 5, 2)
            Widget.setLayout(hLayout)

            return Widget
        load_button = QPushButton(qtawesome.icon('fa.download'),'下载',Widget)
        load_button.setObjectName('Max_button')
        load_button.setCursor(QCursor(Qt.PointingHandCursor))
        load_button.clicked.connect(lambda :self.Download_PPT(name,id))
        show_button = QPushButton(qtawesome.icon('fa.file-image-o'),'浏览', Widget)
        show_button.setObjectName('Max_button')
        show_button.setCursor(QCursor(Qt.PointingHandCursor))
        show_button.clicked.connect(lambda: self.Show_PPT(name,id))
        hLayout = QHBoxLayout()
        hLayout.addWidget(show_button)
        hLayout.addWidget(load_button)
        hLayout.setContentsMargins(5, 2, 5, 2)
        Widget.setLayout(hLayout)
        return Widget
    def Download_PPT(self,name,id):
        self.Request_info('Download_PPT',url_data=self.Request_info('Get_Download_url',url_data=id),Name=name)
    def showImages(self,images_url):
        try:
            self.label.setParent(None)
            self.Qleft.setParent(None)
            self.Qright.setParent(None)
            self.close_button.setParent(None)
            self.gridLayout_images.removeWidget(self.label)
        except:
            pass
        self.label = QtWidgets.QLabel()
        self.label.setObjectName("label")
        self.label.setPixmap(images_url)
        self.label.setScaledContents(True)
        self.close_button = QPushButton(qtawesome.icon('fa.close', color='red'), '', self.label)
        self.close_button.clicked.connect(self.Show_Widget.close)
        self.close_button.setGeometry(680, 0, 30, 30)
        self.close_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.close_button.setStyleSheet('''
                border:none;
                ''')
        self.Qleft = QPushButton(qtawesome.icon('fa.chevron-left',color='gray'), '', self.label)
        self.Qleft.setGeometry(10, 300, 25, 50)
        self.Qleft.setObjectName('Qleft')
        self.Qleft.setIconSize(QtCore.QSize(40, 40))
        self.Qleft.setCursor(QCursor(Qt.PointingHandCursor))
        self.Qright = QPushButton(qtawesome.icon('fa.chevron-right',color='gray'), '', self.label)
        self.Qright.setGeometry(660, 300, 25, 50)
        self.Qright.setIconSize(QtCore.QSize(40, 40))
        self.Qright.setObjectName('Qright')
        self.Qright.setCursor(QCursor(Qt.PointingHandCursor))
        self.Qleft.clicked.connect(lambda: self.Images_Shuffling(self.url_list, 'left'))
        self.Qright.clicked.connect(lambda: self.Images_Shuffling(self.url_list, 'right'))
        # rect = self.label.geometry()
        self.Qright.setStyleSheet('''
        border:none;
        ''')
        self.Qleft.setStyleSheet('''
               border:none;
               ''')
        self.gridLayout_images.addWidget(self.label, 0, 1, 1, 1)
        self.label.setStyleSheet('''  
        
            border-top:6px solid darkGray;
            border-left:6px solid darkGray;
            border-bottom:6px solid darkGray;
            border-right:6px solid darkGray;
            border-top-right-radius:10px;
            border-top-left-radius:10px;
            border-bottom-left-radius:10px;
            border-bottom-right-radius:10px;
        ''')
        # self.gridLayout_images.update()
    def Images_Shuffling(self,Image_list,Type):
        if Type == 'left':
            self.ImagesId -= 1
            if self.ImagesId  < 0: self.ImagesId =  len(Image_list)-1
        elif Type == 'right':
            self.ImagesId += 1
            if self.ImagesId > (len(Image_list)-1) : self.ImagesId = 0
        self.showImages(self.Request_info('Get_Images', Image_list[self.ImagesId]))
    def Mouse_Scroll(self,Tpyt):
        if Tpyt == 'up':
            self.Images_Shuffling(self.url_list, 'left')
        else:
            self.Images_Shuffling(self.url_list, 'right')

    def Show_PPT(self,name,id):
        self.url_list=self.Request_info('show_ppt',url_data=id)
        self.Show_Widget = Show_Imageis()
        self.Show_Widget.setStyleSheet('''
        border-radius: 10px;
        ''')
        self.Show_Widget.resize(600,600)
        self.gridLayout_images = QtWidgets.QGridLayout(self.Show_Widget)
        self.gridLayout_infor = QtWidgets.QGridLayout(self.Show_Widget)
        self.ImagesId=0
        self.showImages(self.Request_info('Get_Images',self.url_list[self.ImagesId]))
        self.Show_Widget.setWindowOpacity(0.9)
        self.Show_Widget.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.Show_Widget.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.Show_Widget.setLayout(self.gridLayout_images)
        self.setLayout(self.gridLayout_infor)
        self.Show_Widget.Signal.connect(self.Mouse_Scroll)
        self.Show_Widget.setStyleSheet('''
        QPushButton{
         border:none;
          font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
          border-radius: 10px;
        }
        QPushButton#Qleft:hover{
         
         border-right:10px solid #00ffffff;
        }
        QPushButton#Qright:hover{border-left:10px solid #00ffffff;}         
                        ''')
        self.Show_Widget.show()
    def PPT_list(self,Page=False):
        if Page:
            body_url= Page
        else:
            self.page_cont = 1
            for data in self.Body_url:
                try:
                    body_url = data[self.sender().text()]
                except:
                    pass
        Details_dict =  self.Request_info('Get_Details',url_data=body_url)
        try:
            self.right_recommend_widget.setParent(None)
            self.tableWidget.setParent(None)
            self.Tpage_lable.setParent(None)
            self.pushButton_down.setParent(None)
            self.pushButton_up.setParent(None)
            self.right_layout.removeWidget(self.right_recommend_widget)
        except:
            pass
        self.tableWidget = QtWidgets.QTableWidget()
        self.right_layout.addWidget(self.tableWidget, 1, 0, 2, 9)
        self.tableWidget.setStyleSheet(
            "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(255, 0, 0, 0), stop:1 rgba(255, 255, 255, 0));\n"
            "gridline-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(255, 0, 0, 0), stop:1 rgba(255, 255, 255, 0));\n"
            "selection-background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(255, 0, 0, 0), stop:1 rgba(255, 255, 255, 0));\n"
            "alternate-background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(255, 0, 0, 0), stop:1 rgba(255, 255, 255, 0));")
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setMinimumSize(100,580)
        self.tableWidget.setFocusPolicy(QtCore.Qt.NoFocus)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.tableWidget.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget.setShowGrid(False)
        self.tableWidget.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(False)

        # self.tableWidget.horizontalHeader().setDefaultSectionSize(170)
        # self.tableWidget.horizontalHeader().setDefaultSectionSize(700)
        self.tableWidget.verticalHeader().setDefaultSectionSize(170)
        self.setStyleSheet('''
        QLabel{
        background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(255, 0, 0, 0), stop:1 rgba(255, 255, 255, 0));
        }
       
        QPushButton#Max_button:hover{border-bottom:7px solid #00EEB422; font-weight:700;}
        QPushButton#Max_button{
        color:#575757;
         font-weight:400;
         border:none;
          font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
        }
        QTableWidget#tableWidget{
            border:none;
            
            font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
        }
                    QScrollBar::handle:horizontal
            {
                background-color: #605F5F;
                min-width: 5px;
                border-radius: 4px;
            }

            QScrollBar::add-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(rc/right_arrow_disabled.png);
                width: 10px;
                height: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(rc/left_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:horizontal:hover,QScrollBar::add-line:horizontal:on
            {
                border-image: url(rc/right_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }


            QScrollBar::sub-line:horizontal:hover, QScrollBar::sub-line:horizontal:on
            {
                border-image: url(rc/left_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:horizontal, QScrollBar::down-arrow:horizontal
            {
                background: none;
            }


            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal
            {
                background: none;
            }

            QScrollBar:vertical
            {
                background-color: #2A2929;
                width: 15px;
                margin: 15px 3px 15px 3px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
            }

            QScrollBar::handle:vertical
            {
                background-color: #605F5F;
                min-height: 5px;
                border-radius: 4px;
            }

            QScrollBar::sub-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(rc/up_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(rc/down_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:vertical:hover,QScrollBar::sub-line:vertical:on
            {

                border-image: url(rc/up_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }


            QScrollBar::add-line:vertical:hover, QScrollBar::add-line:vertical:on
            {
                border-image: url(rc/down_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical
            {
                background: none;
            }


            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical
            {
                background: none;
            }

                    ''')
        self.tableWidget.setColumnCount(3)
        # self.tableWidget.setColumnWidth(2,100)
        self.right_recommend_label.setText("PPT下载列表")
        self.gridLayout_page = QtWidgets.QGridLayout(self.right_widget)
        self.Tpage_lable = QLabel()
        self.pushButton_down = QtWidgets.QPushButton(qtawesome.icon('fa.arrow-down'),'下一页')
        self.pushButton_down.setObjectName('Max_button')
        self.pushButton_down.setCursor(QCursor(Qt.PointingHandCursor))
        self.gridLayout_page.addWidget(self.Tpage_lable, 0, 0, 1, 1)
        self.gridLayout_page.addWidget(self.pushButton_down, 0, 2, 1, 1)
        self.pushButton_up= QtWidgets.QPushButton(qtawesome.icon('fa.arrow-up'),'上一页')
        self.pushButton_up.setObjectName('Max_button')
        self.pushButton_up.setCursor(QCursor(Qt.PointingHandCursor))
        self.gridLayout_page.addWidget(self.pushButton_up, 0, 1, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_page.addItem(spacerItem1, 0, 0, 1, 1)
        self.right_layout.addLayout(self.gridLayout_page, 5, 0, 2, 9)
        if isinstance(Details_dict,dict):
            self.Body_url =Details_dict['kejian']
            # self.tableWidget.setColumnCount(2)
            self.tableWidget.verticalHeader().setDefaultSectionSize(32)
            cont = 0
            cont1 = 0
            for item in Details_dict['kejian']:
                self.pushButton_up.setEnabled(False)
                self.pushButton_down.setEnabled(False)
                Row = self.tableWidget.rowCount()+1
                url_data = list(item.keys())[0]
                if cont == 0:
                    if (Row-1) !=0:cont1+=1
                    self.tableWidget.setRowCount(cont1+1)
                    self.tableWidget.setCellWidget(cont1, 0,self.Download_Button(url_data,item[url_data],True))
                    cont=1
                else:
                    self.tableWidget.setCellWidget(cont1, 1, self.Download_Button(url_data,item[url_data], True))
                    cont = 0
        else:
            self.recommend_label_list = []
            for item in Details_dict:
                self.recommend_label_list.append(QLabel())
                nxet_cont = len(self.recommend_label_list) - 1
                self.recommend_label_list[nxet_cont].setPixmap(self.Request_info('Get_Images',item[list(item.keys())[0]][0]))
                self.recommend_label_list[nxet_cont].setScaledContents(True)
                self.recommend_label_list[nxet_cont].resize(30,30)
                self.recommend_label_list[nxet_cont].setStyleSheet('''
                background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(255, 0, 0, 0), stop:1 rgba(255, 255, 255, 0));
                border-radius: 4px;
                ''')
                Row = self.tableWidget.rowCount() + 1
                self.tableWidget.setRowCount(Row)
                self.tableWidget.setCellWidget(Row - 1, 0,self.recommend_label_list[nxet_cont])
                items = QtWidgets.QTableWidgetItem(list(item.keys())[0])
                items.setTextAlignment(QtCore.Qt.AlignCenter)
                font = QtGui.QFont()
                font.setPointSize(9)
                font.setBold(True)
                font.setWeight(75)
                items.setFont(font)
                self.tableWidget.setItem(Row - 1, 1, items)
                self.tableWidget.setCellWidget(Row - 1, 2,self.Download_Button(list(item.keys())[0],item[list(item.keys())[0]][1]))

            self.page_total_cont = self.Request_info('Get_page', url_data=body_url)

            self.Tpage_lable.setText('总共：%s 页 当前为 %s页' % (self.page_total_cont, self.page_cont))
            self.pushButton_down.clicked.connect(lambda: self.page(body_url))
            self.pushButton_up.clicked.connect(lambda: self.page(body_url))
    def page(self,pag_url):
        if self.sender().text() == '下一页':
            self.page_cont += 1
            if self.page_cont > int(self.page_total_cont): self.page_cont = 1
        elif self.sender().text() == '上一页':
            self.page_cont -= 1
            if self.page_cont < 1:self.page_cont=int(self.page_total_cont)
        self.Tpage_lable.setText('总共：%s 页 当前为 %s页' % (self.page_total_cont, self.page_cont))
        if pag_url.split('/')[-1]:
            pag_url = '/'.join(pag_url.split('/')[0:-1])+'/'
        Pag_Url = '%sppt_%s_%s.html' % (pag_url, (pag_url.split('/'))[-2], self.page_cont)

        if self.Menu_page_kj: Pag_Url='%skejian_%s.html'%(pag_url,self.page_cont)
        if self.Menu_page_sc: Pag_Url='%s%s'%(pag_url,self.page_cont)
        self.PPT_list(Page=Pag_Url)
    def AuthorInformation(self):
        self.AuthorWidget =QWidget()
        self.AuthorWidget.setObjectName('AuthorWidget')
        ClosButton = QHead_Button(b'\xef\x81\xb2'.decode("utf-8"), self.AuthorWidget)
        ClosButton.setObjectName("ClosButton")
        ClosButton.setCursor(QCursor(Qt.PointingHandCursor))
        ClosButton.move(360, 0)
        ClosButton.clicked.connect(self.AuthorWidget.close)
        Qurl_bk=QPushButton('https://fennbk.com',self.AuthorWidget)
        Qurl_bk.setObjectName('Qurl_if')
        Qurl_bk.clicked.connect(lambda :QtGui.QDesktopServices.openUrl(QtCore.QUrl('https://fennbk.com')))
        Qurl_bk.setGeometry(30, 180, 300, 30)
        Qurl_bk.setCursor(QCursor(Qt.PointingHandCursor))
        self.Qurl_git=QPushButton(self.AuthorWidget)
        self.Qurl_git.setObjectName('Qurl_if')
        self.label_type = QLabel('作者：Fenn',self.AuthorWidget)
        self.label_type .setGeometry(40,140,300,30)
        self.label_bk = QLabel('博客：',self.AuthorWidget)
        self.label_bk.setGeometry(40,180,40,30)
        # self.label_bk.connect(lambad:QtGui.QDesktopServices.openUrl(QtCore.QUrl('https:fennbk.com')))
        self.label_git = QLabel('GitHup：',self.AuthorWidget)
        self.label_git.setGeometry(40, 220, 60, 30)
        self.label_git = QLabel('本程序仅限于学习交流，不得用于商业用途\n否则后果自负！',self. AuthorWidget)
        self.label_git.setObjectName('label_git')
        self.label_git.setGeometry(40, 60, 500, 50)
        self.AuthorWidget.resize(400,300)
        # AuthorWidget.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.AuthorWidget.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.gridLayout_infor.addWidget(self.AuthorWidget)
        self.AuthorWidget.setStyleSheet('''
        QPushButton#Qurl_if{color:#0000CD}
        QPushButton#Qurl_if:hover{color:#EE6A50}
        QWidget#AuthorWidget{
        border-top-left-radius:10px;
            border-bottom-left-radius:10px;
            border-bottom-radius:10px;
            
        background-color: qlineargradient(spread:pad, x1:0.524597, y1:1, x2:0.524, y2:0, stop:0 rgba(255, 255, 255, 255), stop:0.0384615 rgba(216, 255, 249, 255), stop:0.387987 rgba(157, 248, 255, 255), stop:0.631494 rgba(156, 237, 255, 255), stop:1 rgba(112, 192, 255, 255));
        }
        QLabel#label_git{color:#EE6A50}
        QLabel{
        background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(255, 0, 0, 0), stop:1 rgba(255, 255, 255, 0));
        }
        ''')
        self.AuthorWidget.show()
    def closeEvent(self, *args, **kwargs):
        try:
            self.AuthorWidget.close()
        except:
            pass
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyleSheet('''
        QPushButton#ClosButton:hover{
            color:#FF6347
        }
        QPushButton#MinMaxButton:hover {
            color:#FF6347
        }
        QPushButton
{

    border-color: #31363b;
    border-radius:50%;
    color: #FFFFFF;


    border-width: 0px;

    border-style: solid;
    padding: 5px;
    border-radius: 2px;
    outline: none;
}
        ''')
    gui = PPT_App()
    gui.show()
    sys.exit(app.exec_())
