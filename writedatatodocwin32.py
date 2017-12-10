#-*- coding:utf-8 -*-
#说明：因为docx不能打包，所以改为win32com的方式实现
import os  
import win32com 
from win32com.client import Dispatch, constants 

import time

import readmbtiresult as rmr
#import dataprocess

#设置Range对象的内容和格式
def SetRangeTextAndFormat(Range,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag):
    Range.Font.Name = fontname
    Range.Font.Size = fontsize
    Range.Font.Italic = itaflag#是否斜体
    Range.Font.Bold = boldflag#是否粗体
    Range.ParagraphFormat.Alignment = alignmentflag # 012左中右
    Range.ParagraphFormat.LeftIndent = leftindent#设置段落格式左缩进
    Range.ParagraphFormat.FirstLineIndent = firstlfind#首行缩进
    
    Range.InsertBefore(txt) # 插入内容
    
    #Range.Text = txt

#设置段落内容和格式
def SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag):
    '''
    功能：设置段落内容和格式
    '''
    p = document.Paragraphs.Add()
    SetRangeTextAndFormat(p.Range,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    return p
    '''
    p.Range.Font.Name = fontname
    p.Range.Font.Size = fontsize
    p.Range.Font.Italic = itaflag#是否斜体
    p.Range.Font.Bold = boldflag#是否粗体
    p.Range.ParagraphFormat.Alignment = alignmentflag # 012左中右
    p.Range.ParagraphFormat.LeftIndent = leftindent#设置段落格式左缩进
    p.Range.ParagraphFormat.FirstLineIndent = firstlfind#首行缩进
    #p.Range.ParagraphFormat.LineSpacing = 12 # 行间距
    p.Range.InsertBefore(txt) # 插入内容
    '''
#设置表格的内容和格式
def SetTableTextAndFormat(cell,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag):
    SetRangeTextAndFormat(cell.Range,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        
def WriteFrontCoverPage(document,StuInfoData):#写封面页
    '''
    功能：写报告的封面
    '''
    ##先添加图像（公司logo图片为logo.png）
    curospath = os.path.abspath('.')
    picpath = curospath + r'\pictures\logo.png'
    p = document.Paragraphs.Add()
    p.Range.InlineShapes.AddPicture(picpath, False,True)

    #两行空行
    SetParagraphTextAndFormat(document,'',16,u'黑体',0,0,0,0,0)
    
    #居中写产品名
    txt = u'北京展梦学业规划指导中心'
    SetParagraphTextAndFormat(document,txt,16,u'黑体',1,0,0,0,0)

    #下面再输入1行空格
    SetParagraphTextAndFormat(document,'',22,u'黑体',1,0,0,0,0)
    
    #居中写报告名
    txt = u'基于MBTI模型职业测评分析报告书'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)

    #下面再输入1行空格
    SetParagraphTextAndFormat(document,'\n',22,u'黑体',1,0,0,0,0)
    '''
    #下面居中输入报告的内容简介    
    txt = u'本报告分为如下几个部分：'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,120,0,0,0)
    txt = u'（1）专业取向分析及建议'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,120,0,0,0)
    txt = u'（2）院校定位及建议'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,120,0,0,0)
    txt = u'（3）附件：自主招生内参资料'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,120,0,0,0)
    '''
    #插入两行空行
    SetParagraphTextAndFormat(document,'\n\n',12,u'宋体',0,120,0,0,0)

    #输出学生信息
    
    name = StuInfoData['name'].decode('utf-8')
    txt = u'姓  名：' + name
    SetParagraphTextAndFormat(document,txt,13,u'黑体',0,120,0,0,0)
    school = StuInfoData['school'].decode('utf-8')
    txt = u'学  校：' + school
    SetParagraphTextAndFormat(document,txt,13,u'黑体',0,120,0,0,0)
    grade = StuInfoData['grade'].decode('utf-8')
    txt = u'年  级：' + grade
    SetParagraphTextAndFormat(document,txt,13,u'黑体',0,120,0,0,0)
    curtime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    txt = u'时  间：' + curtime
    SetParagraphTextAndFormat(document,txt,13,u'黑体',0,120,0,0,0)
    

    #插入3行空行
    SetParagraphTextAndFormat(document,'\n\n\n',12,u'宋体',0,120,0,0,0)
    #输入备注信息：
    txt = u'本报告由北京展梦学业规划指导中心提供，报告内容仅对本次采集的数据负责，如有问题请致电400-88888888' + \
    u'或发送邮件至学业邮箱：xf700@qq.com联系我们！'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,0,25,0,0)

    #插入3行空行
    SetParagraphTextAndFormat(document,'\n\n\n',12,u'宋体',0,120,0,0,0)
        
    #祝福语
    txt = u'北京展梦学业规划指导中心祝莘莘学子心想事成，金榜题名！'
    SetParagraphTextAndFormat(document,txt,12,u'楷体',1,0,0,1,0)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
      
      
#写报告第一页，即报告阅读指南
def WriteReadingGuides(document):#写第一页
    '''
    功能：写报告正文第一页，包括个人信息和专业建议结果
    '''
    #一级标题，个人信息及专业建议
    txt = u'\n基于MBTI模型的职业测评分析报告\n'
    SetParagraphTextAndFormat(document,txt,24,u'黑体',1,0,0,0,0)

    #居中写表格名
    txt = u'\n报告阅读指南\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,False,False)
    
    #报告目的
    txt = u'报告目的'
    SetParagraphTextAndFormat(document,txt,16,u'黑体',0,0,25,False,False)
    
    txt = u'本报告旨在帮助您开始了解和分析最真实的自己，协助您迈出职业定位和职业规划的第一步，' + \
        u'从人格类型的角度描述您的适合岗位特质并您的发展提供建议。'
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,25,False,False)
    
    #报告阅读说明
    txt = u'\n报告阅读说明'
    SetParagraphTextAndFormat(document,txt,16,u'黑体',0,0,25,False,False)
    
    txt = u'（1）本报告对你的人格特点进行了详细描述，它能够帮助你拓展思路，接受更多的可能性，而不是限制你的选择；\n' + \
        u'（2）报告结果（即性格类型）没有“好”与“坏”之分，但不同特点对于不同的工作存在“适合”与“不适合”的区别，从而表现出具体条件下的优、劣势；\n' + \
        u'（3）你的人格特点由遗传、成长环境和生活经历决定，不要想象去改变它，但是我们可以在了解性格的基础上对某些趋向的补充和平衡，从而优化我们的决策，更有效地发挥潜力；\n' + \
        u'（4）报告展示的是你的性格偏好，而不是你的知识、经验、技巧。'
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,25,False,False)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    return document
    
#写报告正文第一部分，包括个人信息和专业建议结果
def WriteCharacterAnalysisInfo(document,MBTITypeLevel,MBTIScore,MBTIDescription):#写第一页
    '''
    功能：写报告正文第一部分，包括个人信息和专业建议结果
    '''
    #前面先加一个空行
    SetParagraphTextAndFormat(document,'',12,u'宋体',0,0,0,0,0)
    #一级标题，人格及MBTI人格理论
    txt = u'人格及MBTI人格理论'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    
    #小标题
    txt = u'\n关于人格\n'
    SetParagraphTextAndFormat(document,txt,20,u'黑体',1,0,0,0,0)
    
    #内容
    txt = u'人格 (personality) 源于拉丁语Persona，也叫个性。心理学中，人格指一个人在一定情况下所作行为反应的特质，即人们在生活、工作中独特的行为表现，包括思考方式、决策方式等。每一种人格理论都是一个用来解释人格的概念、假设、观点和原则的系统。人格问题是一个非常复杂的问题，如果没有一个理论性的指导框架，我们很容易在理解时迷失方向。\n' + \
        u'人格和我们的生活息息相关。陷入爱河，选择朋友，和同事相处，或者和最神经质的亲戚相处，生活处处都离不开人格的痕迹。那么，到底什么是人格？它和气质、性格和态度有什么差异？人格可以测量吗？\n' + \
        u'人格并非意味着“吸引力”、“魅力”或“风格”。心理学认为，人格是一个人独特的思维、情感和行为模式。换句话说，一个人过去是什么样的人，现在和将来还是什么样的人，这种一贯性就是由其人格所决定的。同时，每个人独特的才智、价值观、期望、爱情、仇恨以及习惯等构成的总和，也就使我们每一个人都与众不同。\n' + \
        u'人格和气质也是不同的。气质是形成个性或人格的“原料”之一，是人格的先天遗传成分，它决定着一个人的反应敏感度、活动水平、心境、可塑性。'
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,25,0,0)
    
    #小标题
    txt = u'\n\n人格特质\n'
    SetParagraphTextAndFormat(document,txt,20,u'黑体',1,0,0,0,0)
    
    #内容
    txt = u'心理学家把人的特殊的、稳定的个性品质称为“人格特质”。其实，我们每天都在使用“特质”的概念谈论熟人或朋友。比如，你说你的一个朋友善于交际、办事有条理、聪明。我说我的姐姐是个腼腆、敏感但极有创造型的人。我们所说的这些就是人格特质，是人们在大多数情境下表现出来的稳定的特点，是从观察到的行为中推论出来的。\n' + \
        u'同时，我们经常使用特质预测未来行为。例如，你看到你的朋友总是“见人自来熟”，不论在超市里还是在Party上，与陌生人一谈就说得热热闹闹，由此你可以推论出他具有“善于交际”的特点。然后，你可能会以此为依据，预测他将来工作中也是个爱交际的人。\n' + \
        u'一般人到了二十岁之后，人格就很难再改变了。到了30岁，人格彻底稳定下来，你30岁怎样，到了60岁还差不多是这样。'
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,25,0,0)

    #小标题
    txt = u'\nMBTI人格理论\n'
    SetParagraphTextAndFormat(document,txt,20,u'黑体',1,0,0,0,0)
    
    #内容
    txt = u'1942年，瑞士精神分析学家荣格（弗洛伊德的学生），第一次提出人格分类的概念。他认为感知和判断是大脑的两大基本功能。不同的人，感知倾向不同——有些人更侧重直觉，有些更侧重实感。同样，不同的人判断倾向也不同——有些更倾向理性分析得出结论，有些更侧重情感考量，更为感性。同时，这两大基本功能又受到精力来源不同（内向或外向）的影响。近代美国心理学家Katherine Cook Briggs提出大脑的两大基本功能还受到生活方式倾向的影响：计划型和散漫型。这样，MBTI人格模型便出来了：分为四个维度，每个维度有两个方向，共计八个方面，即共有八种人格特点：'
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,25,0,0)
    
    #写表格名字
    txt = u'表1 MBTI理论人格分类特点'
    SetParagraphTextAndFormat(document,txt,10,u'微软雅黑',1,0,0,0,0)
    
    
    #插入表格
    p = document.Paragraphs.Add()
    table = document.Tables.Add(p.Range, 9, 3)   # 新增一个10*7表格
    #设置表格为网格型
    #if table.Style <> u"网格型":
    table.Style = u"网格型"
    #设置部分表格的高度
    '''
    table.Cell(7,1).Height = 200
    table.Cell(8,1).Height = 60
    table.Cell(9,1).Height = 80
    table.Cell(10,1).Height = 100
    '''
    #设置部分表格的宽度
    for ir in range(1,10):
        table.Cell(ir,1).Width = 40
        table.Cell(ir,2).Width = 80
        table.Cell(ir,3).Width = 300
    '''
    table.Rows.Add()     # 新增一個Row
    table.Columns.Add()     # 新增一個Column
    '''
    #合并单元格
    cell = table.Cell(8,1)
    cell.Merge(table.Cell(9,1))#合并做事方式
    cell = table.Cell(6,1)
    cell.Merge(table.Cell(7,1))#合并决策方式
    cell = table.Cell(4,1)
    cell.Merge(table.Cell(5,1))#合并获取信息的方式 
    cell = table.Cell(2,1)
    cell.Merge(table.Cell(3,1))#合并与世界相互作用方式
    
   
    #表格字体
    fontsize = 10
    fontname = u'微软雅黑'
    alignmentflag = 0
    leftindent = 0
    firstlfind = 0
    itaflag = False 
    boldflag = False
    #表格结构内容
    txt = u'维度'
    SetTableTextAndFormat(table.Cell(1,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,1)
    txt = u'方向'
    SetTableTextAndFormat(table.Cell(1,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,1)
    txt = u'解释'
    SetTableTextAndFormat(table.Cell(1,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,1)
    txt = u'我们与世界相互作用方式'#个人信息
    SetTableTextAndFormat(table.Cell(2,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    table.Cell(2,1).Range.Orientation = 1
    txt = u'我们获取信息的主要方式'
    SetTableTextAndFormat(table.Cell(4,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    table.Cell(4,1).Range.Orientation = 1
    txt = u'我们的决策方式'
    SetTableTextAndFormat(table.Cell(6,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    table.Cell(6,1).Range.Orientation = 1
    txt = u'我们的做事方式'
    SetTableTextAndFormat(table.Cell(8,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    table.Cell(8,1).Range.Orientation = 1
    
    txt = u'外向(E)\nExtraversion'
    SetTableTextAndFormat(table.Cell(2,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'关注自己如何影响外部环境：将心理能量和注意力聚集于外部世界和与他人的交往上。\n例如：聚会、讨论、聊天'
    SetTableTextAndFormat(table.Cell(2,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'内向(I)\nIntroversion'
    SetTableTextAndFormat(table.Cell(3,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'关注外部环境的变化对自己的影响：将心理能量和注意力聚集于内部世界，注重自己的内心体验。\n例如：独立思考，看书，避免成为注意的中心，听的比说的多'
    SetTableTextAndFormat(table.Cell(3,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'感觉(S)\nSensing'
    SetTableTextAndFormat(table.Cell(4,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'关注由感觉器官获取的具体信息：看到的、听到的、闻到的、尝到的、触摸到的事物。\n例如：关注细节、喜欢描述、喜欢使用和琢磨已知的技能'
    SetTableTextAndFormat(table.Cell(4,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'直觉(N)\nIntuition'
    SetTableTextAndFormat(table.Cell(5,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'关注事物的整体和发展变化趋势：灵感、预测、暗示，重视推理。\n例如：重视想象力和独创力，喜欢学习新技能，但容易厌倦、喜欢使用比喻，跳跃性地展现事实'
    SetTableTextAndFormat(table.Cell(5,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'思考(T)\nThinking'
    SetTableTextAndFormat(table.Cell(6,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'重视事物之间的逻辑关系，喜欢通过客观分析作决定评价。\n例如：理智、客观、公正、认为圆通比坦率更重要'
    SetTableTextAndFormat(table.Cell(6,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'情感(F)\nFeeling'
    SetTableTextAndFormat(table.Cell(7,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'以自己和他人的感受为重，将价值观作为判定标准。\n例如：有同情心、善良、和睦、善解人意，考虑行为对他人情感的影响，认为圆通和坦率同样重要'
    SetTableTextAndFormat(table.Cell(7,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'判断(J)\nJudging'
    SetTableTextAndFormat(table.Cell(8,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'喜欢做计划和决定，愿意进行管理和控制，希望生活井然有序。\n例如：重视结果（重点在于完成任务）、按部就班、有条理、尊重时间期限、喜欢做决定'
    SetTableTextAndFormat(table.Cell(8,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'知觉(P)\nPerceiving'
    SetTableTextAndFormat(table.Cell(9,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'灵活、试图去理解、适应环境、倾向于留有余地，喜欢宽松自由的生活方式。\n例如：重视过程、随信息的变化不断调整目标，喜欢有多种选择。'
    SetTableTextAndFormat(table.Cell(9,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)

    #内容
    txt = u'在以上四个维度上，每个人都会有自己天生就具有的倾向性，也就是说，处在两个方向分界点的这边或那边，我们称之为“偏好”。例如：如果你落在外向的那边，称为“你具有外向的偏好”；如果你落在内向的那边，称为“你具有内向的偏好”。\n' + \
        u'在现实生活中，每个维度的两个方面你都会用到，只是其中的一个方面你用的更频繁、更舒适，就好像每个人都会用到左手和右手，习惯用左手的人是左撇子，习惯用右手的人是右撇子。同样，在四个维度上你用的最频繁、最熟练的那种方式就是你在这个维度上的偏好，而这四个偏好加以组合，就形成了你的人格类型，它反映了你在一系列心理过程和行为方式上的特点。我们不仅拥有人格特质，还有每个人所独有的才智、价值观、期望、爱情、仇恨以及习惯，这些使得我们每一个人都与众不同。'
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,25,0,0)
    
    ##############################################################################
    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #一级标题#你的MBTI性格
    txt = u'\n您的MBTI性格类型\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    
    #内容
    txt = u'下面的表格及其后面的几段文字提供您所报告的有关您的性格类型的解读。您所指示的四种性格倾向中的每一种都有对应的分数，分数越高，您所表达的那种性格倾向就越明显。'
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,25,0,0)
    
    #表2 原始得分及性格倾向明显程度表
    #写表格名字
    txt = u'表2 原始得分及性格倾向明显程度表'
    SetParagraphTextAndFormat(document,txt,10,u'微软雅黑',1,0,0,0,0)
    
    p = document.Paragraphs.Add()
    table = document.Tables.Add(p.Range, 5, 3)   # 新增一个10*7表格
    table.Style = u"网格型"
    for ir in range(1,6):
        table.Cell(ir,1).Width = 120
        table.Cell(ir,2).Width = 180
        table.Cell(ir,3).Width = 120
    #表格字体
    fontsize = 10
    fontname = u'微软雅黑'
    alignmentflag = 0
    leftindent = 0
    firstlfind = 0
    itaflag = False 
    boldflag = False
    #表格结构内容
    txt = u'维度'
    SetTableTextAndFormat(table.Cell(1,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,1)
    txt = u'得分'
    SetTableTextAndFormat(table.Cell(1,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,1)
    txt = u'倾向明显程度'
    SetTableTextAndFormat(table.Cell(1,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,1)
    txt = u'与世界相互作用方式'#个人信息
    SetTableTextAndFormat(table.Cell(2,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'获取信息方式'
    SetTableTextAndFormat(table.Cell(3,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'决策方式'
    SetTableTextAndFormat(table.Cell(4,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'做事方式'
    SetTableTextAndFormat(table.Cell(5,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'外向（E）/内向（I）=' + str(MBTIScore[0]) + '/' + str(MBTIScore[1])
    SetTableTextAndFormat(table.Cell(2,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'感觉（S）/直觉（N）=' + str(MBTIScore[2]) + '/' + str(MBTIScore[3])
    SetTableTextAndFormat(table.Cell(3,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'思考（T）/情感（F）=' + str(MBTIScore[4]) + '/' + str(MBTIScore[5])
    SetTableTextAndFormat(table.Cell(4,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'判断（J）/知觉（P）=' + str(MBTIScore[6]) + '/' + str(MBTIScore[7])
    SetTableTextAndFormat(table.Cell(5,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = MBTITypeLevel[0].decode('utf-8')
    SetTableTextAndFormat(table.Cell(2,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = MBTITypeLevel[1].decode('utf-8')
    SetTableTextAndFormat(table.Cell(3,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = MBTITypeLevel[2].decode('utf-8')
    SetTableTextAndFormat(table.Cell(4,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = MBTITypeLevel[3].decode('utf-8')
    SetTableTextAndFormat(table.Cell(5,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    
    #结论
    txt = u'\n测试显示，您的MBTI性格类型为：' + MBTIDescription['type'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,14,u'微软雅黑',0,0,25,0,1)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #写性格描述信息
    #一级标题，您的个性特征描述
    txt = u'\n您的个性特征描述\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    #内容
    txt = MBTIDescription['characterdescription'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,0,0,0)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #一级标题，可能存在的盲点
    txt = u'\n可能存在的盲点\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    #内容
    txt = MBTIDescription['possibleblindspot'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,0,0,0)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #一级标题，如何有效地使用你的类型
    txt = u'\n如何有效地使用你的类型\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    #内容
    txt = MBTIDescription['howtouse'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,0,0,0)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #一级标题，工作中的优势
    txt = u'\n工作中的优势\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    #内容
    txt = MBTIDescription['advantage'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,0,0,0)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #一级标题，工作中的劣势
    txt = u'\n工作中的劣势\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    #内容
    txt = MBTIDescription['disadvantage'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,0,0,0)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #一级标题，适合的岗位特质
    txt = u'\n适合的岗位特质\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    #内容
    txt = MBTIDescription['suitablepost'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,0,0,0)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #一级标题，适合的职业类型
    txt = u'\n您适合的职业\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    #内容
    txt = MBTIDescription['suitableprofession'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,0,0,0)

    #插入分页
    p = document.Paragraphs.Add()    
    p.Range.InsertBreak()
    
    #一级标题，个人发展建议
    txt = u'\n个人发展建议\n'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)
    #内容
    txt = MBTIDescription['suggestions'].decode('utf-8')
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,0,0,0)

    '''
    #写在最后：
    txt = u'\n\n\n由衷的感谢您选择北京展梦学业规划指导中心来为您规划学业，'
    txt += u'同时也感谢您耐心的填写并提交了MBTI职业测评，为了您能够更加准备和全面的了解'
    txt += u'自己的性格特征及所适合的职业类型，我们向您提供该分析报告，'
    txt += u'如对报告内容有疑问，请按照封面上的联系方式'
    txt += u'联系我们的学业规划老师邵老师进行详细的咨询，谢谢！'
    SetParagraphTextAndFormat(document,txt,12,u'微软雅黑',0,0,25,0,0)
    '''
    
    return document

#将信息写入word文件总调用函数
def WriteMBTITestResult2Doc(FilePathStu,iStu):
    '''
    功能：写doc文件总调用函数
    '''
    curospath = os.path.abspath('.')
    FilePathMBTIDescription = curospath + '//data//MBTIDescriptionData.xlsx'

    flag,MBTIDescriptionData = rmr.ReadMBTIDescriptionData(FilePathMBTIDescription)
    if flag == 0:
        return 0
    #读取不同学生和学校的报考信息
    flag,StuData,AnswerList = rmr.ReadStuMBTITestResult(FilePathStu)
    if flag == 1:  
        stunum = len(StuData)
    else:
        return 0

    LoopList = []
    if len(iStu) == 1:
        if iStu[0] >= 0 and iStu[0] < stunum:
            LoopList.append(iStu[0])
        else:
            LoopList = range(stunum)
    else:
        #检查有效性
        for i in range(len(iStu)):
            if iStu[i] < 0 or iStu[i] >= stunum:
                del iStu[i]
        LoopList = iStu
        
    #word写入准备
    #一次打开word引擎
    #打开word引擎
    w = win32com.client.Dispatch('Word.Application') 
    # 后台运行，不显示，不警告
    w.Visible = 0
    w.DisplayAlerts = 0 

    itype = -1#第几种性格类型
    for i in LoopList:#循环写入
        MBTIType,MBTITypeLevel,MBTIScore = rmr.GetMBTIType(AnswerList[i])#根据测试结果判断MBTI类型及分数
        for it in range(16):
            if MBTIType in MBTIDescriptionData[it]['type']:
                itype = it
        
        document = w.Documents.Add() # 创建新的文档，对于每个学生创建一个新文档
        WriteFrontCoverPage(document,StuData[i])#写封面页
        WriteReadingGuides(document)#写报告指南
        if itype >= 0 and itype <= 15:
            WriteCharacterAnalysisInfo(document,MBTITypeLevel,MBTIScore,MBTIDescriptionData[itype])#写学校的信息

        docpath = curospath + '\\报告\\' + StuData[i]['name'] + '.docx'#str(i) + '.docx'#
        docpath = docpath.decode('utf-8')
        document.SaveAs(docpath)
        document.Close()

    w.Quit()
    
    return 1


if __name__=="__main__":
    FilePathTestResult = r'E:\xfedu_workstation\MBTITestDoc\data\MBTI2017-12-08.xlsx'
    iStu = [0]
    WriteMBTITestResult2Doc(FilePathTestResult,iStu)

