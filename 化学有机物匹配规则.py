import re
from functools import cmp_to_key

#化学物质俗名列表 其中@字符后面加数字代表匹配顺序，数字小的排在前面
gOrgTrivialNames = """2,4-滴@-1 青蒿(素|醇|烯|酸|酮) 香草(素|醇|酸|酯|醛) 橙花醇 
利尿酸 白藜芦醇 香兰素 达菲 莽草酸 雌烯醇 甲基(橙|红) 小檗(胺|碱) 谷甾(醇|烷) 蜜胺 金鸡纳(霜|素|碱) 丁香 巴豆 育亨 
败脂酸 香树精 天冬 萜品醇 法尼 安息香 土霉素 谷胱甘肽 神经氨酸 青霉胺 呋喃坦啶 凤梨(醛|酯) 组胺 奎宁 蚁酸 
门冬酰胺 紫(尿|脲)酸 紫霉素 (黄|磺)原酸 雌(二|酚|酮|醇|甾)  四环素 唑嘧胺 四氢巴马亭 多巴 延胡索乙素 棕榈 
锥虫胂胺 周效磺胺 仲丁通 中草酸 正定霉素 长春碱 长春新碱 脂肪 樟柳碱 粘液酸 粘菌素 增甘(磷|膦) 增塑剂 增效(醚|胺)
增血压素 增产灵 云杉苷 晕苯 晕海宁 月桂 鸢尾 异烟 异长叶烷酮 麻黄碱 软脂酸 卡必醇 哈尔碱 加兰他敏 柠檬酸 月桂酸 
麦角含硫碱 高胱氨酸 咪唑啉 噻吩 噻嗪 呋喃 吡啶 吡喃 葡萄糖 黄体酮 孕甾烷 雄甾烷 雌甾烷 胆烷 胆甾烷 麦角甾烷 
豆甾烷 吡咯 糠酰胺 冰?醋酸 鞘氨醇 甘油 龙脑 冰片 糠醛 育亨烷 吗啡 半日花烷  古罗糖酸 胆固醇 胡萝卜素 膦咛 枯烯 
百里香素 富烯 噁烷 呋咱 咪唑 吡嗪 嘧啶 哒嗪 吗啉 色胺 哌啶 哌嗪 昴苯 锥苯 卵苯 色烯 吲哚 嘌呤 喹嗪 喹啉 蝶啶 
色烷 吲哚啉 奎宁啶 葡萄庚糖 吖啶 吩嗪 螺旋烃 呫吨 喹喔啉 色酚 葡糖醛 硬脂酸 苹果酸 水杨(醛|酸|酰|醇|胺) 脯氨酸 
黄酮 樟脑 核糖 喹哪啶 腺嘌呤 核甙酸 腺甙酸 胞甙酸 尿甙酸 鸟嘌呤 鸟甙酸 胞嘧啶 甲状腺素 环氧树脂 皮质甾酮 吡唑啉 
吡唑 喹吖啶 青霉素 蝶呤 噻唑 甘氨酸 焦袂康酸 肉桂 大茴香 巴豆醛 茴香 安替比林 哇啉 香豆(素|酸|醛) 儿茶酚 松油醇 
萜品烯醇 雄甾烯 愈创木 肌甙酸 鸟苷酸 巴比妥酸 紫罗兰(酮|醇|炔|酯) 紫罗酮 茶碱 半乳糖 高半胱氨酸 阿拉伯糖 甘露糖 
酒石酸 来苏糖? 唾液酸 环香叶醇 酪氨酸 荧光素 泼尼松 琥珀酸 青霉烷 谷氨酸 萄葡糖 天门冬氨酸 木糖 糠基 咔唑 酚酞  
草酸 炔雌醇 香茅醛 乳酸 桂酸 桂醛 吡哆 涕丙酸 马来酸 氨基酸 色氨酸 牛磺酸 组氨酸 茶氨酸 薄荷醇 薄荷酯 扁桃酸 
茉莉酮 卡巴 甘醇 蓖酸 油酸 泛酸 果糖 荧烷 乳糖 蔗糖 丁子香 睾丸素 天仙子 木质素 睾丸酮 纤维素 尿素 氯吡格雷
咖啡(因|酸|醛|酮|醇|酰) 贝诺酯 避蚊胺 苯佐卡因 桂皮酸 苯嗪草酮 尿酸 醋硝香豆素 阿司匹林 蒿甲醚 溴氰菊酯"""
#有机物可能的歧义字符串，例如：乙醇氧化，乙醇氧不作为有机物
gAmbiguousWords = """过氧化氢 另一 之一 一代 二代 三代 四代 提高 通过 经过 做过 盛过 放过 
推断 判断 碱石灰 酸性 氧化 酯化 金属 无色 氧原子 脱氢反应 根据 过高 石墨烯 氢离子 环境 , 一单元 苯环 氢氧化钠 
羟基氢 羟基氧 分开 次氯酸 (一) (二) (三) (四) (五) (六) (七) (八) (九) (1) (2) (3) (4) (5) (6) (7) (8) (9)
(酚) 溴水"""
#判断是否是有机物
gOrgJudge = """胺 菲 蒽 蕃 醌 芘 脲 肟 薁 噻 嗪 茚 萘 苄 苊 芴 苉 苝 荭 蔻 莒 苣 脒 
苷 腙 肼 膦 砜 胍 胂 萜 茴 蒈 莰 吡 芪 蒎 唑 苯 茋 胼 (.+酚)|(酚.+) (.+酰)|(酰.+) (.+羟)|(羟.+) 
(.+羰)|(羰.+) (.+羧)|(羧.+) (.+酸酐)|(酸酐.+) (.+醇)|(醇.+) (.+酯)|(酯.+) (.+醚)|(醚.+) 
(.+烷)|(烷.+) (.+烯)|(烯.+) (.+炔)|(炔.+) (.+醛)|(醛.+) (.+酮)|(酮.+) (甲|乙|丙|丁|戊|己|庚|辛|壬|癸).*酸  
(一|二|三|四|五|六|七|八|九|十|廿|杂|双).*环  (一|二|三|四|五|六|七|八|九|十|廿|多).*元 .+糖 .+橡胶  氰"""
#有机物可能的组成项，这些项不出现在头部和尾部
gOrgNotHeadTail = """- — ﹣ － 0 化 代 替 杂 合 并 缩 内 叉基 叉 撑 去 元 """
#有机物可能的组成项，这些项不出现在尾部
gOrgNotTail = """ ( [ { 一 二 三 四 五 六 七 八 九 十 廿 
\(*([0-9]|[a-z]|[A-Z]|[α-ω]|\,|，|'|′|’|\(|\)|\[|反|对|\|])+(-|－|─|—|—|﹣|­-)\)* 磺 硝 甲 乙 丙 丁 戊 己 已  
庚 辛 壬 癸 单 双 叁 肆 伍 陆 柒 捌 玖 拾 联 聚 脱 失 降 增 高 断 开 迁 移 逆 正 异 新 伯 仲 叔 季 顺式 顺 反式 
反 映 邻 间位 间 对称 对,对’ 对氨基苯甲酸 对- 对二甲苯 对叔丁基 对亚硝基 对乙基 对异丙基 迫 聚 亚 双 偶 特 卤 重 
甘露 过 全 多 硬脂 不对称 左旋 连 偏 均 共聚 氢化 水合 次 盐酸 环"""
#有机物可能的组成项，这些项不出现在头部
gOrgNotHead = """ ) ] } 酰 树脂 树酯 巯 基 叉基 亚基 爪基 次基 自由基 根 酸酐 酸 酐 
烷 桥 烃 正离子 负离子 离子 共聚物 粘合剂 锂 钠 钾 铷 铯 钫 铍 镁 钙 锶 钡 镭 铝 镓 铟 铊 锡 铅 砷 锑 铋 硒 
碲 钋 砹 氦 氖 氩 氪 氙 氡 钪 钛 钒 铬 锰 钴 镍 铜 锌 钇 锆 铌 钼 锝 钌 铑 钯 银 镉 铪 钽 钨 铼 锇 铱 铂 
金 汞 镧 铈 镨 钕 钷 钐 铕 钆 铽 镝 钬 铒 铥 镱 镥 锕 钍 镤 铀 镎 钚 镅 锔 锫 锎 锿 镄 钔 锘 铹 共聚乳液 
化物 胶乳 糖 宾 碱 """
#有机物可能的组成项
gOrgElements = """胺 菲 蒽 蕃 醌 芘 脲 腈 肟 薁 噻 嗪 茚 萘 苄 苊 芴 苉 苝 荭 蔻 芬 
莒 苣 吩 脒 腙 肼 膦 砜 吡 胍 酞 胂 茂 萜 茴 蒈 莰 芪 蒎 唑 氨 苯 氰 氟 氯 溴 碘 茋 酚 酮 碳 硅 锗 啶 羟 醛 
醇 酯 醚 羰 羧 炔 烯 硼 氢 硫 氮 磷 氧 铁 铵 甙 脎 胲 尿 苷 胼 诺 糠 枯 荧 盐 嘧 橡胶 胸腺 纤维
 \[(\d|\,|，|[α-ω])*\]"""
gIsInited = False

class ChemistryOrganicMatch:
    class OrganicItem:
        def __init__(self, match, st, len):
            self.mMatch = match
            self.mIndex = st
            self.mLength = len
        def __str__(self):
            return self.mMatch.mText[self.mIndex : self.mIndex + self.mLength]
        def __getitem__(self, offset):
            return self.mMatch.mText[self.mIndex + offset]
        
    mText = ""
    mOrganicItems = []
    def __init__(self):
        self.mOrganicItems = []
        pass
    def __init__(self, text):
        self.mOrganicItems = []
        self.ReMatch(text)
    def __iter__(self):
        return iter(self.mOrganicItems)
    def __len__(self):
        return len(self.mOrganicItems)
    def ReMatch(self,text):
        self.mText = text
        self.Match()
    #//匹配字符串
    def RemoveAmbiguous(self, organic):
        if organic.mLength == 0:
            return
        matched = True
        while matched:
            global mAmbiguous
            for amb in mAmbiguous:
                matched = False
                #//扩展字符串
                start = organic.mIndex - len(amb) + 1;
                if start < 0 or start >= len(self.mText):
                    start = 0
                end = organic.mIndex + organic.mLength + len(amb) - 1;
                if end > len(self.mText) or end < 0:
                    end = len(self.mText)
                #//扩展后的匹配结果
                ext = self.mText[start: end];
                #//左边的歧义位置
                left = ext.find(amb);
                #//不存在歧义字符串
                if left < 0 or left + start > organic.mIndex:
                    left = organic.mIndex;
                #//歧义位置位于开头处则抛弃歧义字符串
                else:
                    matched = True;
                    left = left + start + len(amb);
                #//右侧的歧义字符串位置
                right = ext.rfind(amb);
                #//不存在歧义字符串
                if (right <= 0 or right + start + len(amb) < organic.mIndex + organic.mLength):
                    right = organic.mIndex + organic.mLength;
                #//歧义位置位于结尾处则抛弃歧义字符串
                else:
                    matched = True;
                    right = right + start
                if (right < left):
                    right = left
                #//记录去除歧义后的结果
                organic.mIndex = left; organic.mLength = right - left;
                if len(str(organic)) == 0:
                    break

    def Match(self):
        self.InitOnlyFirst()
        global mAllRegex
        for match in re.finditer(mAllRegex, self.mText):
            if match.start() == match.end():
                continue
            organic = self.OrganicItem(self, match.start(), match.end() - match.start());
            #去除歧义,去除开头和结尾的歧义字符串
            self.RemoveAmbiguous(organic)
            #处理开头，去除不能作为开头的特征词
            global mNotHead
            nhmatch = re.match(mNotHead, str(organic))
            while nhmatch is not None:
                #//存在不能作为开头的特征词，则不断去除
                if nhmatch.start() == 0 and nhmatch.end() != nhmatch.start():
                    organic.mIndex += nhmatch.end() - nhmatch.start()
                    organic.mLength -= nhmatch.end() - nhmatch.start()
                    nhmatch = re.match(mNotHead, str(organic))
                #//开头特征词符合要求，则跳出
                else:
                    break
            # 处理结尾
            results = []
            #//获取所有不能作为结尾的特征词
            global mNotTail
            for matchNT in re.finditer(mNotTail,str(organic)):
                results.append(self.OrganicItem(self, organic.mIndex + matchNT.start(), matchNT.end() - matchNT.start()));
            #//反转
            results = reversed(results)
            #//从最后一个开始查找
            for ite in results:
                #//如果存在，则去除结尾特征词
                if (ite.mIndex + ite.mLength) == (organic.mIndex + organic.mLength):
                    organic.mLength -= ite.mLength;
                #//不存在，则跳出
                else:
                    break
            #region 括号匹配，去除开始和末尾不匹配的括号
            stackl = []; stackr = []
            for i in range(0, organic.mLength):
                if (organic[i] == '(' or organic[i] == '['):
                    stackl.append(i)
                elif(organic[i] == ')' or organic[i] == ']'):
                    if (len(stackl) > 0):
                        stackl.pop()
                    #//右括号不匹配
                    else:
                        stackr.append(i);
                    
            #//去除首尾不匹配的括号
            lTrip = 0; rTrip = organic.mLength - 1;
            while lTrip in stackl or lTrip in stackr:
                lTrip += 1;
            while rTrip in stackl or rTrip in stackr:
                rTrip -= 1;
            organic.mIndex += lTrip;
            organic.mLength = rTrip - lTrip + 1;
            if organic.mLength > 0 and organic[0] == '(' and organic[organic.mLength - 1] == ')':
                organic.mIndex += 1
                organic.mLength -= 2
            self.RemoveAmbiguous(organic)
            #region 必须存在的有机标志
            global mMustExist
            existmatch = re.search(mMustExist, str(organic))
            if existmatch:
                self.mOrganicItems.append(organic)
        
    def InitOnlyFirst(self):
        global gIsInited
        if not gIsInited:
            gIsInited = True
            #//歧义词
            global mAmbiguous
            mAmbiguous = gAmbiguousWords.split()
            ##//去除空字符串
            #mAmbiguous.RemoveAll(x => string.IsNullOrWhiteSpace(x));
            #//所有有机物组成成分
            global mAllRegex 
            mAllRegex = "%s %s %s %s %s"%( gOrgTrivialNames, gOrgElements, gOrgNotHead, gOrgNotTail, gOrgNotHeadTail)
            #//生成匹配正则表达式
            mAllRegex = self.ToRegex(mAllRegex);
            global mNotHead
            mNotHead = self.ToRegex("%s %s"%( gOrgNotHead, gOrgNotHeadTail))
            global mNotTail 
            mNotTail = self.ToRegex("%s %s"%( gOrgNotTail, gOrgNotHeadTail))
            global mMustExist 
            mMustExist = self.ToRegex("%s %s"%( gOrgTrivialNames, gOrgJudge))
    def RegPunctuation(self, qs):
        qs = qs.replace(r" , ", r" \, ")
        qs = qs.replace(r" . ", r" \. ")
        qs = qs.replace(r" : ", r" \: ")
        qs = qs.replace(r" ( ", r" \( ")
        qs = qs.replace(r" ) ", r" \) ")
        qs = qs.replace(r" [ ", r" \[ ")
        qs = qs.replace(r" ] ", r" \] ")
        qs = qs.replace(r" { ", r" \{ ")
        qs = qs.replace(r" } ", r" \} ")
        qs = qs.replace(r" ' ", r" ' ")
        qs = qs.replace(r" * ", r" \* ")
        qs = qs.replace(r" - ", r" \- ")
        return qs;

    #//将以空格分割的名称转化为正则表达式
    def ToRegex(self, allRegex):
        #//特殊符号转义
        allRegex = self.RegPunctuation(allRegex)
        #//获取所有的匹配列表
        allRegxLst = allRegex.split()
        ##//去除空字符串
        #allRegxLst.RemoveAll(x => string.IsNullOrWhiteSpace(x));
        #//去除重复
        allRegxLst = list(set(allRegxLst))
        #//排序，如果没有指定序号，则长的放前面
        allRegxLst.sort(key=cmp_to_key(itemCmp))
                
        #//生成正则匹配项
        allRegex = "|".join(allRegxLst)
        #//去除序号标志@-1
        allRegex = re.sub(r"(@-{0,1}[0-9]+)", "",allRegex)
        #//生成正则表达式
        return "(" + allRegex + ")+"

def itemCmp(l,r):
    #//@-1 @0 @1 匹配项后面的匹配顺序标志，默认为0，按序从前向后匹配，如果匹配顺序相同，则按字典排序
    il = 0 if '@' not in l else int(l[l.index('@') + 1:])
    ir = 0 if '@' not in r else int(r[r.index('@') + 1:])
    #//序号小的排前面
    if il != ir:
        return il - ir
    else:
        #//长的排前面，相同的按字典排序
        if len(l) == len(r):
            if l > r:
                return 1
            elif l == r:
                return 0
            else:
                return -1 
        else:
            return len(r)- len(l)