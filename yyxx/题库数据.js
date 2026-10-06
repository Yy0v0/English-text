// ═══════════════════════════════════════════════════════════════════
//  题库数据文件 · 星耀学院
//  本文件为游戏题库的唯一真实数据源（Source of Truth）
//
//  【重要】修改题库的方法：
//  ① 编辑本文件 → 保存 → 直接刷新游戏页面（无需运行任何脚本）
//     推荐！最简单，直接改直接用。
//
//  ② 通过 Excel 编辑 → 运行「同步题库.py」→ 生成新版「题库数据.js」
//     适合多人协作编辑 xlsx，再同步到游戏。
//
//  数据格式说明（每条记录）：
//  { type, diff, q, options, ans, explain }
//  type : '单词' | '语法' | '翻译'
//  diff : 'low'(小学低) | 'high'(小学高) | 'chu1'(初一) | 'chu2'(初二)
//  q    : 题目文字
//  options: [选项A, 选项B, 选项C, 选项D]  （数组，顺序与ans对应）
//  ans  : 正确答案索引，0=A, 1=B, 2=C, 3=D
//  explain: 答对后显示的解析文字
// ═══════════════════════════════════════════════════════════════════

const QUESTION_BANK = [
  // ========== 小学低年级 (low) ==========
  // --- 单词 ---
  {type:'单词', diff:'low', q:'「苹果」的英文是？',        options:['apple','orange','banana','grape'],        ans:0, explain:'apple = 苹果'},
  {type:'单词', diff:'low', q:'「猫」的英文是？',          options:['dog','bird','cat','fish'],               ans:2, explain:'cat = 猫'},
  {type:'单词', diff:'low', q:'「书」的英文是？',          options:['pen','book','bag','desk'],                ans:1, explain:'book = 书'},
  {type:'单词', diff:'low', q:'「学校」的英文是？',        options:['hospital','school','park','store'],        ans:1, explain:'school = 学校'},
  {type:'单词', diff:'low', q:'「红色」的英文是？',        options:['blue','green','red','yellow'],            ans:2, explain:'red = 红色'},
  {type:'单词', diff:'low', q:'「星期一」的英文是？',      options:['Sunday','Monday','Tuesday','Friday'],     ans:1, explain:'Monday = 星期一'},
  {type:'单词', diff:'low', q:'「水」的英文是？',          options:['milk','juice','water','tea'],             ans:2, explain:'water = 水'},
  {type:'单词', diff:'low', q:'「大」的英文是？',          options:['small','tall','big','short'],             ans:2, explain:'big = 大'},
  {type:'单词', diff:'low', q:'「家」的英文是？',          options:['room','house','home','door'],             ans:2, explain:'home / house = 家'},
  {type:'单词', diff:'low', q:'「快乐」的英文是？',        options:['happy','sad','angry','tired'],             ans:0, explain:'happy = 快乐'},
  {type:'单词', diff:'low', q:'「老师」的英文是？',        options:['student','doctor','teacher','farmer'],    ans:2, explain:'teacher = 老师'},
  {type:'单词', diff:'low', q:'「花园」的英文是？',        options:['garden','kitchen','bedroom','bathroom'],  ans:0, explain:'garden = 花园'},
  {type:'单词', diff:'low', q:'「下雨」的英文是？',        options:['snow','rain','wind','cloud'],             ans:1, explain:'rain = 下雨'},
  {type:'单词', diff:'low', q:'「聪明」的英文是？',        options:['clever','brave','kind','lazy'],           ans:0, explain:'clever = 聪明'},
  {type:'单词', diff:'low', q:'「男孩」的英文是？',        options:['girl','boy','woman','child'],             ans:1, explain:'boy = 男孩'},
  {type:'单词', diff:'low', q:'「女孩」的英文是？',        options:['boy','girl','man','woman'],               ans:1, explain:'girl = 女孩'},
  {type:'单词', diff:'low', q:'「狗」的英文是？',          options:['cat','dog','bird','fish'],                ans:1, explain:'dog = 狗'},
  {type:'单词', diff:'low', q:'「鱼」的英文是？',          options:['cat','bird','fish','dog'],                ans:2, explain:'fish = 鱼'},
  {type:'单词', diff:'low', q:'「小」的英文是？',          options:['big','small','tall','long'],              ans:1, explain:'small = 小'},
  {type:'单词', diff:'low', q:'「高」的英文是？',          options:['short','tall','small','long'],            ans:1, explain:'tall = 高'},
  // --- 语法 ---
  {type:'语法', diff:'low', q:'选出正确的句子：',          options:['I am a student.','I is a student.','I are student.','Me am student.'],       ans:0, explain:'主语I配am，a student冠词不能省'},
  {type:'语法', diff:'low', q:'「她喜欢吃苹果」英文正确的是？', options:['She like apples.','She likes apple.','She likes apples.','She is like apples.'], ans:2, explain:'第三人称单数，like加-s，apple用复数'},
  {type:'语法', diff:'low', q:'「这是一本书」英文正确的是？',  options:['This is a book.','This is book.','That is a book.','This are a book.'],     ans:0, explain:'This is + a/an + 单数名词'},
  {type:'语法', diff:'low', q:'「他是一个男孩」英文正确的是？',  options:['He is a boy.','He are a boy.','She is a boy.','He is boy.'],             ans:0, explain:'He = 他，用is，a boy表类别'},
  {type:'语法', diff:'low', q:'选出正确的句子：',          options:['They are student.','They are students.','They is students.','They am students.'], ans:1, explain:'They是复数，be动词用are，名词用复数'},
  // --- 翻译 ---
  {type:'翻译', diff:'low', q:'"Good morning!"的意思是？',      options:['晚安','你好','早上好','再见'],                       ans:2, explain:'Good morning = 早上好'},
  {type:'翻译', diff:'low', q:'"How old are you?"意思是？',     options:['你叫什么名字？','你在哪里？','你多大了？','你怎么了？'],       ans:2, explain:'How old = 多大/几岁'},
  {type:'翻译', diff:'low', q:'"Thank you very much!"意思是？', options:['对不起','非常感谢','不客气','再见'],                   ans:1, explain:'Thank you = 谢谢，very much = 非常'},
  {type:'翻译', diff:'low', q:'"What time is it?"意思是？',    options:['几点了？','在哪里？','什么东西？','多少钱？'],              ans:0, explain:'What time = 什么时间/几点'},
  {type:'翻译', diff:'low', q:'"I love you!"意思是？',        options:['我恨你','我爱我','我认识你','我帮你'],                   ans:1, explain:'love = 爱'},
  {type:'翻译', diff:'low', q:'"See you tomorrow!"意思是？',   options:['昨天见','明天见','今天见','再见'],                       ans:1, explain:'See you tomorrow = 明天见'},
  {type:'翻译', diff:'low', q:'"What\'s your name?"意思是？', options:['你好吗？','你叫什么？','你在做什么？','你多大了？'],           ans:1, explain:'What\'s your name = 你叫什么名字'},
  {type:'翻译', diff:'low', q:'"I\'m hungry."意思是？',      options:['我很开心','我很饿','我很累','我很忙'],                     ans:1, explain:'hungry = 饿的'},
  {type:'翻译', diff:'low', q:'"How do you do?"意思是？',    options:['你怎么样？','你好！','你做什么？','你去哪？'],                 ans:1, explain:'How do you do? 是正式问候语'},
  {type:'翻译', diff:'low', q:'"It\'s raining outside."意思是？', options:['外面在下雨','外面在刮风','外面很热','外面很冷'],         ans:0, explain:'raining = 在下雨'},
  {type:'翻译', diff:'low', q:'"Good night!"意思是？',       options:['早安','午安','晚安','明天见'],                           ans:2, explain:'Good night = 晚安（睡前道别）'},
  {type:'翻译', diff:'low', q:'"Nice to meet you!"意思是？',  options:['你好！','很高兴认识你！','再见！','谢谢！'],               ans:1, explain:'Nice to meet you = 很高兴认识你'},

  // ========== 小学高年级 (high) ==========
  // --- 单词 ---
  {type:'单词', diff:'high', q:'「医生」的英文是？',         options:['nurse','doctor','teacher','driver'],              ans:1, explain:'doctor = 医生'},
  {type:'单词', diff:'high', q:'「图书馆」的英文是？',        options:['cinema','library','museum','park'],               ans:1, explain:'library = 图书馆'},
  {type:'单词', diff:'high', q:'「夏天」的英文是？',         options:['spring','summer','autumn','winter'],              ans:1, explain:'summer = 夏天'},
  {type:'单词', diff:'high', q:'「生日」的英文是？',         options:['party','birthday','holiday','festival'],           ans:1, explain:'birthday = 生日'},
  {type:'单词', diff:'high', q:'「朋友」的英文是？',         options:['family','friend','classmate','teacher'],           ans:1, explain:'friend = 朋友'},
  {type:'单词', diff:'high', q:'「电脑」的英文是？',         options:['television','computer','telephone','radio'],        ans:1, explain:'computer = 电脑'},
  {type:'单词', diff:'high', q:'「医院」的英文是？',         options:['hospital','school','cinema','museum'],              ans:0, explain:'hospital = 医院'},
  {type:'单词', diff:'high', q:'「练习」的英文是？',         options:['practice','homework','exercise','test'],            ans:2, explain:'exercise = 练习'},
  {type:'单词', diff:'high', q:'「蔬菜」的英文是？',         options:['fruit','vegetable','meat','bread'],                ans:1, explain:'vegetable = 蔬菜'},
  {type:'单词', diff:'high', q:'「骑自行车」的英文是？',     options:['ride a bike','drive a car','fly a kite','run fast'], ans:0, explain:'ride a bike = 骑自行车'},
  {type:'单词', diff:'high', q:'「帮助」的英文是？',         options:['help','hope','hold','hurt'],                       ans:0, explain:'help = 帮助'},
  {type:'单词', diff:'high', q:'「厨房」的英文是？',         options:['kitchen','garden','bedroom','bathroom'],           ans:0, explain:'kitchen = 厨房'},
  {type:'单词', diff:'high', q:'「春天」的英文是？',         options:['spring','summer','autumn','winter'],              ans:0, explain:'spring = 春天'},
  {type:'单词', diff:'high', q:'「秋天」的英文是？',         options:['spring','summer','autumn','winter'],              ans:2, explain:'autumn = 秋天'},
  {type:'单词', diff:'high', q:'「圣诞节」的英文是？',       options:['Christmas','Halloween','Easter','Valentine'],     ans:0, explain:'Christmas = 圣诞节'},
  // --- 语法 ---
  {type:'语法', diff:'high', q:'"There ___ a book on the table."',      options:['are','is','be','were'],              ans:1, explain:'There is/are：单数用is，复数用are'},
  {type:'语法', diff:'high', q:'"I ___ my homework yesterday."',        options:['do','does','did','doing'],           ans:2, explain:'yesterday表示过去时，用did'},
  {type:'语法', diff:'high', q:'"She can ___ English well."',          options:['speaks','speaking','spoke','speak'], ans:3, explain:'情态动词can后接动词原形'},
  {type:'语法', diff:'high', q:'"He ___ to school every day."',        options:['go','goes','going','went'],           ans:1, explain:'一般现在时，第三人称单数动词加-es'},
  {type:'语法', diff:'high', q:'"They ___ playing football now."',    options:['is','are','was','am'],                ans:1, explain:'They是复数，现在进行时用are + doing'},
  {type:'语法', diff:'high', q:'选出正确的句子：',                    options:['He don\'t like fish.','He doesn\'t likes fish.','He doesn\'t like fish.','He not like fish.'], ans:2, explain:'第三人称否定用doesn\'t，后接动词原形'},
  {type:'语法', diff:'high', q:'"___ you like some tea?"',              options:['Do','Would','Are','Have'],            ans:1, explain:'Would you like...? = 你想要...吗？'},
  {type:'语法', diff:'high', q:'"The cat is ___ than the dog."',        options:['small','smaller','smallest','more small'], ans:1, explain:'small的比较级是smaller，不用more'},
  {type:'语法', diff:'high', q:'"Let\'s ___ football."',               options:['play','plays','playing','to play'],   ans:0, explain:'Let\'s = 让我们，后接动词原形'},
  {type:'语法', diff:'high', q:'"I have ___ apple."',                  options:['a','an','the','/'],                   ans:1, explain:'apple以元音音素开头，用an'},
  // --- 翻译 ---
  {type:'翻译', diff:'high', q:'"How are you?"意思是？',             options:['你叫什么？','你几岁？','你好吗？','你在哪？'],                  ans:2, explain:'How are you = 你好吗？'},
  {type:'翻译', diff:'high', q:'"Where is the hospital?"意思是？',    options:['医院在哪？','医院是什么？','医院很大','我喜欢医院'],            ans:0, explain:'Where is = ...在哪里？'},
  {type:'翻译', diff:'high', q:'"I was born in 2015."意思是？',       options:['我住在2015年','我出生于2015年','我出生于2014年','我在2015年去世'], ans:1, explain:'be born in = 出生于'},
  {type:'翻译', diff:'high', q:'"She often reads books."意思是？',   options:['她有时读书','她经常读书','她从不读书','她喜欢读书'],            ans:1, explain:'often = 经常，reads = 读书（第三人称单数）'},
  {type:'翻译', diff:'high', q:'"There are 50 students in our class."意思是？', options:['我们班有50个学生','我们班是50个学生','我们班有15个学生','我们班很棒'], ans:0, explain:'There are ... = 有...（数量）'},
  {type:'翻译', diff:'high', q:'"My favorite subject is English."意思是？', options:['我不喜欢英语','我最喜欢的科目是英语','英语很难','我喜欢数学'], ans:1, explain:'favorite subject = 最喜欢的科目'},
  {type:'翻译', diff:'high', q:'"Don\'t be late!"意思是？',          options:['请进来','别迟到！','请坐','别说话'],                          ans:1, explain:'Don\'t + 动词原形 = 别...'},
  {type:'翻译', diff:'high', q:'"She looks like her mother."意思是？',  options:['她不喜欢她妈妈','她看起来像她妈妈','她在家','她很漂亮'],       ans:1, explain:'looks like = 看起来像'},

  // ========== 初一 (chu1) ==========
  // --- 单词 ---
  {type:'单词', diff:'chu1', q:'「环境」的英文是？',           options:['environment','industry','culture','society'],     ans:0, explain:'environment = 环境'},
  {type:'单词', diff:'chu1', q:'「国际」的英文是？',           options:['national','international','social','personal'],    ans:1, explain:'international = 国际的'},
  {type:'单词', diff:'chu1', q:'「发展」的英文是？',           options:['develop','developing','development','developed'],  ans:0, explain:'develop = 发展（动词）'},
  {type:'单词', diff:'chu1', q:'「保护」的英文是？',           options:['protect','prevent','provide','produce'],          ans:0, explain:'protect = 保护'},
  {type:'单词', diff:'chu1', q:'「传统」的英文是？',           options:['modern','traditional','popular','natural'],      ans:1, explain:'traditional = 传统的'},
  {type:'单词', diff:'chu1', q:'「表演」的英文是？',           options:['performance','perform','form','transform'],       ans:1, explain:'perform = 表演（动词）'},
  {type:'单词', diff:'chu1', q:'「经历」的英文是？',           options:['experience','experiment','expert','explain'],    ans:0, explain:'experience = 经历（名词/动词）'},
  {type:'单词', diff:'chu1', q:'「节日」的英文是？',           options:['festival','feast','food','culture'],             ans:0, explain:'festival = 节日'},
  {type:'单词', diff:'chu1', q:'「分享」的英文是？',           options:['share','spare','square','smart'],                ans:0, explain:'share = 分享'},
  {type:'单词', diff:'chu1', q:'「传统节日」是哪个词？',        options:['traditional food','traditional festival','modern festival','popular festival'], ans:1, explain:'traditional festival = 传统节日'},
  // --- 语法 ---
  {type:'语法', diff:'chu1', q:'"If you ___ help, call me."',        options:['need','needs','needed','needing'],       ans:0, explain:'If条件句中用一般现在时，need为动词原形'},
  {type:'语法', diff:'chu1', q:'"I have lived here ___ 2010."',       options:['for','since','in','at'],                 ans:1, explain:'since + 过去时间点，表示从...开始'},
  {type:'语法', diff:'chu1', q:'"The book ___ by my father."',       options:['was written','is writing','writes','written'], ans:0, explain:'被动语态：be + 过去分词'},
  {type:'语法', diff:'chu1', q:'"I\'d like ___ tea, please."',       options:['some','any','many','much'],               ans:0, explain:'would like / want后用some（表示请求）'},
  {type:'语法', diff:'chu1', q:'"___ the window, please. It\'s cold."', options:['Close','Open','Clean','Break'],       ans:0, explain:'根据语境"It\'s cold"应关窗，祈使句用动词原形'},
  {type:'语法', diff:'chu1', q:'"She asked me ___ there."',          options:['to go','goes','going','went'],             ans:0, explain:'ask sb. to do sth. = 让某人做某事'},
  {type:'语法', diff:'chu1', q:'"It\'s very important ___ healthy food."', options:['eat','to eat','eating','eaten'],   ans:1, explain:'It\'s important to do sth. = 做某事很重要'},
  {type:'语法', diff:'chu1', q:'"He has ___ finished his homework."',  options:['yet','already','still','also'],         ans:1, explain:'already用于肯定句，表示已经'},
  // --- 翻译 ---
  {type:'翻译', diff:'chu1', q:'"English is used widely around the world."意思是？', options:['英语只在英国使用','英语在世界各地广泛使用','英语很难学','英语很有趣'], ans:1, explain:'be used widely = 被广泛使用'},
  {type:'翻译', diff:'chu1', q:'"We should protect the environment."意思是？',        options:['我们应该保护环境','我们破坏了环境','我们喜欢环境','我们需要环境'], ans:0, explain:'should protect = 应该保护'},
  {type:'翻译', diff:'chu1', q:'"The festival celebrates the harvest."意思是？',     options:['节日庆祝旅游','节日庆祝丰收','节日庆祝工作','节日庆祝学习'], ans:1, explain:'celebrates = 庆祝，harvest = 丰收'},
  {type:'翻译', diff:'chu1', q:'"I think traditional culture is important."意思是？', options:['我认为传统文化不重要','我认为传统文化很重要','我不了解传统文化','我讨厌传统文化'], ans:1, explain:'I think + 句子 = 我认为...，important = 重要'},
  {type:'翻译', diff:'chu1', q:'"Learning English helps us understand the world."意思是？', options:['学英语没什么用','学英语帮助我们了解世界','世界很难理解','我们不理解英语'], ans:1, explain:'Learning English = 学英语，helps us = 帮助我们'},

  // ========== 初二 (chu2) ==========
  // --- 单词 ---
  {type:'单词', diff:'chu2', q:'「环境问题」的英文是？',          options:['environment problem','environmental problems','problem environment','environment issues'], ans:1, explain:'environmental problems = 环境问题'},
  {type:'单词', diff:'chu2', q:'「污染」的英文是？',              options:['protect','pollution','population','position'],                        ans:1, explain:'pollution = 污染'},
  {type:'单词', diff:'chu2', q:'「全球变暖」的英文是？',          options:['global warming','world hot','earth temperature','climate cold'],          ans:0, explain:'global warming = 全球变暖'},
  {type:'单词', diff:'chu2', q:'「可持续的」的英文是？',          options:['sustainable','available','reliable','suitable'],                       ans:0, explain:'sustainable = 可持续的'},
  {type:'单词', diff:'chu2', q:'「影响」的英文是？',              options:['infect','affect','effect','effort'],                                  ans:1, explain:'affect = 影响（动词）'},
  {type:'单词', diff:'chu2', q:'「解决方案」的英文是？',          options:['solution','pollution','population','situation'],                     ans:0, explain:'solution = 解决方案'},
  {type:'单词', diff:'chu2', q:'「减少」的英文是？',              options:['increase','reduce','produce','recycle'],                             ans:1, explain:'reduce = 减少'},
  {type:'单词', diff:'chu2', q:'「回收」的英文是？',              options:['recycle','reduce','reuse','reflect'],                                  ans:0, explain:'recycle = 回收利用'},
  {type:'单词', diff:'chu2', q:'「气候变化」的英文是？',          options:['weather change','climate change','temperature change','season change'],  ans:1, explain:'climate change = 气候变化'},
  {type:'单词', diff:'chu2', q:'「自然保护区」的英文是？',      options:['nature reserve','natural reserve','reserve nature','protect nature'],     ans:0, explain:'nature reserve = 自然保护区'},
  {type:'单词', diff:'chu2', q:'「塑料袋」的英文是？',            options:['plastic bag','poly bag','shop bag','use bag'],                        ans:0, explain:'plastic bag = 塑料袋'},
  // --- 语法 ---
  {type:'语法', diff:'chu2', q:'"___ we protect nature, we will have a better future."', options:['If','Unless','When','Because'],     ans:0, explain:'If条件句：If + 现在时，will + 动词原形'},
  {type:'语法', diff:'chu2', q:'"He asked me where I ___."',                                 options:['live','lived','living','lives'],  ans:1, explain:'间接引语中过去时态后退一步'},
  {type:'语法', diff:'chu2', q:'"If you don\'t reduce pollution, the earth ___ hotter."',   options:['will get','gets','got','getting'], ans:0, explain:'If not + 现在时，will + 动词原形（真实条件句）'},
  {type:'语法', diff:'chu2', q:'"She told me she ___ to Beijing the next day."',            options:['would go','will go','goes','went'], ans:0, explain:'过去将来时：would + 动词原形'},
  {type:'语法', diff:'chu2', q:'"More and more people ___ to use public transport."',      options:['begin','began','have begun','begins'], ans:2, explain:'现在完成时，表示已经发生的动作'},
  {type:'语法', diff:'chu2', q:'"We should take action ___ climate change."',               options:['stop','to stop','stopping','stopped'], ans:1, explain:'take action to do sth. = 采取行动做某事'},
  {type:'语法', diff:'chu2', q:'"Not only ___ but also she sings beautifully."',          options:['she dances','does she dance','she dance','she dancing'], ans:1, explain:'Not only开头的倒装：Not only + 助动词 + 主语'},
  {type:'语法', diff:'chu2', q:'"___ using plastic bags, we can protect the environment."', options:['Stop','Stopping','By stopping','Stopped'], ans:2, explain:'By + doing = 通过...方式'},
  {type:'语法', diff:'chu2', q:'"If everyone ___ an environmentalist, the earth ___ much better."', options:['is','are','was','am'], ans:0, explain:'Everyone is = 每个人都是'},
  // --- 翻译 ---
  {type:'翻译', diff:'chu2', q:'"Climate change is a serious problem."意思是？',      options:['气候变化是个小问题','气候变化是个严重的问题','气候变化不是问题','气候变化很容易解决'], ans:1, explain:'serious problem = 严重的问题'},
  {type:'翻译', diff:'chu2', q:'"We should reduce waste and protect nature."意思是？', options:['我们应该增加垃圾','我们应该减少浪费，保护自然','我们应该破坏环境','我们应该忽视自然'], ans:1, explain:'reduce waste = 减少浪费，protect nature = 保护自然'},
  {type:'翻译', diff:'chu2', q:'"Renewable energy helps solve pollution problems."意思是？', options:['新能源会造成污染','可再生能源有助于解决污染问题','污染不是问题','能源不能解决问题'], ans:1, explain:'renewable energy = 可再生能源，solve = 解决'},
  {type:'翻译', diff:'chu2', q:'"Everyone can make a difference to the environment."意思是？', options:['没人能改变环境','每个人都能对环境产生影响','环境不受影响','只有政府能改变环境'], ans:1, explain:'make a difference = 产生影响'},
  {type:'翻译', diff:'chu2', q:'"Rising temperatures cause many problems."意思是？',   options:['温度下降造成问题','温度上升造成许多问题','温度没有问题','温度引起关注'], ans:1, explain:'rising temperatures = 上升的温度，cause = 造成'},
  {type:'翻译', diff:'chu2', q:'"Taking public transport is a green way to travel."意思是？', options:['乘坐公共交通是一种绿色出行方式','公共交通很慢','绿色不重要','出行很贵'], ans:0, explain:'green way = 绿色方式'},
];
